use anyhow::{bail, Context, Result};
use chrono::Local;
use pytools_core::archive_ops::{
    archive_by_date_copy, archive_by_date_preview, archive_by_type_copy, archive_by_type_preview,
    execute_archive_plan,
};
use pytools_core::batch_ops::{
    build_quarterly_rural_loan_plan, build_remove_foreign_plan, build_rural_loan_plan,
    build_village_split_plan,
};
use pytools_core::batch_rename::run_batch_rename as run_batch_rename_business;
use pytools_core::config::{load_extract_tasks, load_global, GlobalConfig};
use pytools_core::config_admin::{build_config_repair_plan, initialize_new_config, view_config};
use pytools_core::conversion::{
    convert_excel_files_in_source_folders, convert_excel_files_to, convert_one,
    convert_word_files_in_source_folders, convert_word_files_to,
};
use pytools_core::dedup_ops::{
    build_append_by_config_plans, build_comment_duplicate_plan, build_config_duplicate_plan,
    precheck_dedup_tasks, write_new_append_candidate, AppendTargetPlan, DuplicateAction,
};
use pytools_core::edit_plan::{apply_plan_file, apply_plan_with_capability_file};
use pytools_core::file_commit::{
    copy_workbook_to_new_target, create_target_from_candidate, replace_target_from_candidate,
    sha256_file, verify_planning_snapshot,
};
use pytools_core::file_sort_prefix::{
    run_content_prefix_rename_xlsx_candidate, run_content_prefix_unrename_xlsx_candidate,
    run_file_sort_prefix as run_file_sort_prefix_business,
    run_remove_file_sort_prefix as run_remove_file_sort_prefix_business,
};
use pytools_core::local_filename_map::run_local_filename_map_candidate;
use pytools_core::print_ops::{
    build_comment_print_plan, build_config_print_plans, precheck_print_tasks,
};
use pytools_core::quick::{open_config, open_latest_output, open_output_directory};
use pytools_core::region_sum::run_region_sum_check;
use pytools_core::sheet_extract::run_extract_sheets_candidate;
use pytools_core::sheet_rename::{build_sheet_rename_plan, run_batch_sheet_rename_candidate};
use pytools_core::submission_check::run_submission_check;
use pytools_core::summary_ops::{
    summarize_by_comments, summarize_by_comments_with_progress, summarize_used_ranges,
    summarize_used_ranges_with_progress, SummaryProgress, SummaryRunEvent,
};
use pytools_core::timeline_ops::{
    build_wide_target_append_plans, prepare_wide_target_replacement_source,
    run_header_path_compare, run_timeline_precheck, run_wide_summary,
    validate_wide_target_replacement_candidate,
};
use std::fs::{self, File, OpenOptions};
use std::io::{self, BufRead, IsTerminal, Write};
use std::path::{Path, PathBuf};
use std::time::{Instant, SystemTime, UNIX_EPOCH};

fn main() {
    if let Err(error) = run() {
        eprintln!("pytools: {error:#}");
        std::process::exit(2);
    }
}

fn run() -> Result<()> {
    let mut arguments = std::env::args().skip(1);
    let Some(command) = arguments.next() else {
        return run_interactive_menu(Vec::new());
    };
    if command == "--help" || command == "-h" {
        print_usage();
        return Ok(());
    }
    let arguments = arguments.collect::<Vec<_>>();
    if arguments
        .iter()
        .any(|argument| matches!(argument.as_str(), "--help" | "-h"))
    {
        print_usage();
        return Ok(());
    }
    match command.as_str() {
        "menu" | "interactive" => run_interactive_menu(arguments),
        "timeline-precheck" => run_timeline_precheck_cli(arguments),
        "timeline-compare" => run_timeline_compare_cli(arguments),
        "wide-summary" => run_wide_summary_cli(arguments, false),
        "wide-summary-row-suffix" => run_wide_summary_cli(arguments, true),
        "summary-used-range" => run_summary_used_range_cli(arguments),
        "summary-comments" => run_summary_comments_cli(arguments),
        "archive-date-preview" => run_archive_cli(arguments, "date-preview"),
        "archive-type-preview" => run_archive_cli(arguments, "type-preview"),
        "archive-date-copy" => run_archive_cli(arguments, "date-copy"),
        "archive-type-copy" => run_archive_cli(arguments, "type-copy"),
        "archive-execute-plan" => run_archive_cli(arguments, "plan"),
        "convert-excel" => run_convert_batch_cli(arguments, true),
        "convert-word" => run_convert_batch_cli(arguments, false),
        "convert-file" => run_convert_file_cli(arguments),
        "config-init" => run_config_init_cli(arguments),
        "config-repair-plan" => run_config_repair_plan_cli(arguments),
        "config-view" => run_config_view_cli(arguments),
        "dedup-comment-plan" => run_dedup_comment_plan(arguments),
        "dedup-config-plan" => run_dedup_config_plan(arguments),
        "dedup-precheck" => run_dedup_precheck(arguments),
        "dedup-append-plan" => run_dedup_append_plan(arguments),
        "dedup-append-execute" => run_dedup_append_execute(arguments),
        "apply-edit-plan" => run_apply_edit_plan(arguments),
        "print-comment-plan" => run_print_comment_plan(arguments),
        "print-config-precheck" => run_print_config_precheck(arguments),
        "print-config-plan" => run_print_config_plan(arguments),
        "village-split-plan" => run_village_split_plan(arguments),
        "village-split" => run_village_split_batch(arguments),
        "remove-foreign-bank-plan" => run_remove_foreign_bank_plan(arguments),
        "remove-foreign-bank" => run_remove_foreign_bank_batch(arguments),
        "rural-loan-plan" => run_rural_loan_plan(arguments),
        "quarterly-rural-loan-plan" => run_quarterly_rural_loan_plan(arguments),
        "institution-normalize-plan" => run_institution_normalize_plan(arguments),
        "institution-normalize" => run_institution_normalize_batch(arguments),
        "fx-header" => run_fx_header_batch(arguments),
        "fx-header-plan" => run_fx_header_plan(arguments),
        "rename-sheets-plan" => run_rename_sheets_plan(arguments),
        "rename-sheets" => run_rename_sheets(arguments),
        "submission-check" => run_submission_check_cli(arguments),
        "open-config" => run_quick_cli(arguments, "config"),
        "open-latest-output" => run_quick_cli(arguments, "latest"),
        "open-output-dir" => run_quick_cli(arguments, "output"),
        "extract-sheets" | "extract-sheets-area" => run_extract_sheets(arguments),
        "region-sum-check" => run_region_sum(arguments),
        "file-sort-prefix" => run_file_sort_prefix(arguments),
        "remove-file-sort-prefix" => run_remove_file_sort_prefix(arguments),
        "rename-by-content" => run_content_prefix_rename_cli(arguments),
        "unrename-by-content" => run_content_prefix_unrename_cli(arguments),
        "rename-local-map" => run_local_filename_map(arguments),
        "rename-batch" => run_batch_rename(arguments),
        _ => {
            print_usage();
            bail!("unsupported command {command:?}")
        }
    }
}

fn run_interactive_menu(arguments: Vec<String>) -> Result<()> {
    let options = parse_cli_arguments(arguments)?;
    anyhow::ensure!(
        options.positional.is_empty(),
        "menu accepts options only; pass workbook paths after selecting a menu item"
    );
    let config_path = options
        .values
        .get("config")
        .map(PathBuf::from)
        .unwrap_or_else(default_app_config_path);
    let log_view = options
        .values
        .get("log-view")
        .map(String::as_str)
        .unwrap_or("user");
    anyhow::ensure!(
        matches!(log_view, "user" | "developer" | "off"),
        "--log-view must be user, developer, or off"
    );
    let capabilities = options.values.get("capabilities").map(PathBuf::from);
    let platform = options.values.get("platform").cloned();
    anyhow::ensure!(
        capabilities.is_some() == platform.is_some(),
        "--capabilities and --platform must be provided together"
    );
    let global = load_global(&config_path).with_context(|| {
        format!(
            "load PyTools global configuration {}",
            config_path.display()
        )
    })?;

    let stdin = io::stdin();
    let allow_native_dialog = stdin.is_terminal();
    let mut input = stdin.lock();
    let stdout = io::stdout();
    let mut output = stdout.lock();
    run_interactive_main_menu(
        &mut input,
        &mut output,
        &config_path,
        &global,
        log_view,
        allow_native_dialog,
        capabilities.as_deref(),
        platform.as_deref(),
    )
}

fn default_app_config_path() -> PathBuf {
    std::env::current_exe()
        .ok()
        .and_then(|executable| executable.parent().map(Path::to_path_buf))
        .unwrap_or_else(|| PathBuf::from("."))
        .join("config")
        .join("config.xlsx")
}

fn run_interactive_main_menu<R, W>(
    input: &mut R,
    output: &mut W,
    config_path: &Path,
    global: &GlobalConfig,
    log_view: &str,
    allow_native_dialog: bool,
    capabilities: Option<&Path>,
    platform: Option<&str>,
) -> Result<()>
where
    R: BufRead,
    W: Write,
{
    loop {
        writeln!(output, "======== PyTools 菜单 ========")?;
        writeln!(output, "  1  时序工具（已接入 1、3、4、5）")?;
        writeln!(output, "  2  去重工具（检查与生成计划）")?;
        writeln!(output, "  3  打印工具（生成页面设置计划）")?;
        writeln!(output, "  4  转换与文件命名工具")?;
        writeln!(output, "  6  汇总工具")?;
        writeln!(output, "  7  一二批处理工具（当前接入 7-1 至 7-5）")?;
        writeln!(output, "  0  退出")?;
        writeln!(output, "==============================")?;
        let Some(choice) = read_menu_choice(input, output, "主菜单选择> ")? else {
            writeln!(output, "再见。")?;
            return Ok(());
        };
        match choice.as_str() {
            "1" => run_interactive_timeline_menu(input, output, config_path, allow_native_dialog)?,
            "2" => {
                run_interactive_dedup_menu(input, output, config_path, global, allow_native_dialog)?
            }
            "3" => {
                run_interactive_print_menu(input, output, config_path, global, allow_native_dialog)?
            }
            "4" => run_interactive_convert_menu(
                input,
                output,
                config_path,
                global,
                allow_native_dialog,
            )?,
            "6" => {
                run_interactive_summary_menu(input, output, global, log_view, allow_native_dialog)?
            }
            "7" => run_interactive_batch_menu(
                input,
                output,
                config_path,
                global,
                log_view,
                allow_native_dialog,
                capabilities,
                platform,
            )?,
            "0" => {
                writeln!(output, "再见。")?;
                return Ok(());
            }
            _ => writeln!(output, "仅接受 1、2、3、4、6、7 或 0")?,
        }
    }
}

fn config_cli_arguments(config_path: &Path) -> Vec<String> {
    vec![
        "--config".to_owned(),
        config_path.to_string_lossy().into_owned(),
    ]
}

fn run_interactive_timeline_menu<R, W>(
    input: &mut R,
    output: &mut W,
    config_path: &Path,
    allow_native_dialog: bool,
) -> Result<()>
where
    R: BufRead,
    W: Write,
{
    loop {
        writeln!(output, "------ 时序工具 ------")?;
        writeln!(output, "  1  表头路径比对")?;
        writeln!(output, "  3  宽表汇总（行头加后缀）")?;
        writeln!(output, "  4  宽表汇总（行头不加后缀）")?;
        writeln!(output, "  5  预校验")?;
        writeln!(output, "  0  返回")?;
        writeln!(output, "----------------------")?;
        let Some(choice) = read_menu_choice(input, output, "子菜单选择> ")? else {
            return Ok(());
        };
        match choice.as_str() {
            "1" => {
                let Some(template) = choose_workbooks(
                    "选择表头路径比对模板",
                    false,
                    input,
                    output,
                    allow_native_dialog,
                )?
                .and_then(|paths| paths.into_iter().next()) else {
                    writeln!(output, "[已取消]")?;
                    return Ok(());
                };
                let Some(sources) = choose_workbooks(
                    "选择表头路径比对源文件（可多选）",
                    true,
                    input,
                    output,
                    allow_native_dialog,
                )?
                else {
                    writeln!(output, "[已取消]")?;
                    return Ok(());
                };
                let mut arguments = config_cli_arguments(config_path);
                arguments.push("--template".to_owned());
                arguments.push(template.to_string_lossy().into_owned());
                arguments.extend(
                    sources
                        .into_iter()
                        .map(|path| path.to_string_lossy().into_owned()),
                );
                run_timeline_compare_cli(arguments)?;
                return Ok(());
            }
            "3" | "4" => {
                let Some(sources) = choose_workbooks(
                    if choice == "3" {
                        "选择行头加后缀宽表汇总源文件（可多选）"
                    } else {
                        "选择行头不加后缀宽表汇总源文件（可多选）"
                    },
                    true,
                    input,
                    output,
                    allow_native_dialog,
                )?
                else {
                    writeln!(output, "[已取消]")?;
                    return Ok(());
                };
                let mut arguments = config_cli_arguments(config_path);
                arguments.extend(
                    sources
                        .into_iter()
                        .map(|path| path.to_string_lossy().into_owned()),
                );
                run_wide_summary_cli(arguments, choice == "3")?;
                return Ok(());
            }
            "5" => {
                run_timeline_precheck_cli(config_cli_arguments(config_path))?;
                return Ok(());
            }
            "0" => return Ok(()),
            _ => writeln!(output, "仅接受 1、3、4、5 或 0")?,
        }
    }
}

fn run_interactive_dedup_menu<R, W>(
    input: &mut R,
    output: &mut W,
    config_path: &Path,
    global: &GlobalConfig,
    allow_native_dialog: bool,
) -> Result<()>
where
    R: BufRead,
    W: Write,
{
    loop {
        writeln!(output, "------ 去重工具 ------")?;
        writeln!(output, "  1  按批注检查并标记重复（生成计划）")?;
        writeln!(output, "  2  按批注删除重复（生成计划）")?;
        writeln!(output, "  3  按配置检查并标记重复（生成计划）")?;
        writeln!(output, "  4  按配置删除重复（生成计划）")?;
        writeln!(output, "  5  按配置追加到目标（生成计划）")?;
        writeln!(output, "  6  按配置预校验")?;
        writeln!(output, "  0  返回")?;
        writeln!(output, "----------------------")?;
        let Some(choice) = read_menu_choice(input, output, "子菜单选择> ")? else {
            return Ok(());
        };
        match choice.as_str() {
            "1" | "2" => {
                let Some(input_path) = choose_workbooks(
                    if choice == "1" {
                        "选择按批注检查并标记重复的工作簿"
                    } else {
                        "选择按批注删除重复的工作簿"
                    },
                    false,
                    input,
                    output,
                    allow_native_dialog,
                )?
                .and_then(|paths| paths.into_iter().next()) else {
                    writeln!(output, "[已取消]")?;
                    return Ok(());
                };
                let candidate = unique_workbook_candidate(&input_path, &global.output_dir, "dedup");
                let mut arguments = config_cli_arguments(config_path);
                arguments.extend([
                    "--input".to_owned(),
                    input_path.to_string_lossy().into_owned(),
                    "--output".to_owned(),
                    candidate.to_string_lossy().into_owned(),
                    "--action".to_owned(),
                    if choice == "1" {
                        "check".to_owned()
                    } else {
                        "delete".to_owned()
                    },
                ]);
                let action_name = if choice == "1" {
                    "comment-check"
                } else {
                    "comment-delete"
                };
                let action_output_dir = unique_menu_output_dir(&global.output_dir, action_name);
                let plan_path = action_output_dir.join("plan.json");
                arguments.push("--plan-output".to_owned());
                arguments.push(plan_path.to_string_lossy().into_owned());
                run_dedup_comment_plan(arguments)?;
                return Ok(());
            }
            "3" | "4" => {
                let action_name = if choice == "3" {
                    "config-check"
                } else {
                    "config-delete"
                };
                let action_output_dir = unique_menu_output_dir(&global.output_dir, action_name);
                let mut arguments = config_cli_arguments(config_path);
                arguments.extend([
                    "--output-dir".to_owned(),
                    action_output_dir.to_string_lossy().into_owned(),
                    "--action".to_owned(),
                    if choice == "3" {
                        "check".to_owned()
                    } else {
                        "delete".to_owned()
                    },
                ]);
                run_dedup_config_plan(arguments)?;
                return Ok(());
            }
            "5" => {
                let action_output_dir = unique_menu_output_dir(&global.output_dir, "config-append");
                let mut arguments = config_cli_arguments(config_path);
                arguments.extend([
                    "--output-dir".to_owned(),
                    action_output_dir.to_string_lossy().into_owned(),
                ]);
                run_dedup_append_plan(arguments)?;
                return Ok(());
            }
            "6" => {
                run_dedup_precheck(config_cli_arguments(config_path))?;
                return Ok(());
            }
            "0" => return Ok(()),
            _ => writeln!(output, "仅接受 1、2、3、4、5、6 或 0")?,
        }
    }
}

fn unique_menu_output_dir(output_root: &Path, action: &str) -> PathBuf {
    let timestamp = Local::now().format("%Y%m%d_%H%M%S_%f");
    let stem = format!("dedup-{action}-{timestamp}");
    let mut candidate = output_root.join(&stem);
    let mut suffix = 2;
    while candidate.exists() {
        candidate = output_root.join(format!("{stem}-{suffix}"));
        suffix += 1;
    }
    candidate
}

fn unique_menu_action_dir(output_root: &Path, action: &str) -> PathBuf {
    let timestamp = Local::now().format("%Y%m%d_%H%M%S_%f");
    let stem = format!("pytools-{action}-{timestamp}");
    let mut candidate = output_root.join(&stem);
    let mut suffix = 2;
    while candidate.exists() {
        candidate = output_root.join(format!("{stem}-{suffix}"));
        suffix += 1;
    }
    candidate
}

fn run_interactive_print_menu<R, W>(
    input: &mut R,
    output: &mut W,
    config_path: &Path,
    global: &GlobalConfig,
    allow_native_dialog: bool,
) -> Result<()>
where
    R: BufRead,
    W: Write,
{
    loop {
        writeln!(output, "------ 打印工具 ------")?;
        writeln!(output, "  1  按批注保留格式（生成计划）")?;
        writeln!(output, "  2  按批注快速复制（生成计划）")?;
        writeln!(output, "  3  按配置准备打印文件（生成计划）")?;
        writeln!(output, "  4  按配置预校验")?;
        writeln!(output, "  0  返回")?;
        writeln!(output, "----------------------")?;
        let Some(choice) = read_menu_choice(input, output, "子菜单选择> ")? else {
            return Ok(());
        };
        match choice.as_str() {
            "1" | "2" => {
                let Some(input_path) = choose_workbooks(
                    if choice == "1" {
                        "选择按批注保留格式准备打印的工作簿"
                    } else {
                        "选择快速复制准备打印的工作簿"
                    },
                    false,
                    input,
                    output,
                    allow_native_dialog,
                )?
                .and_then(|paths| paths.into_iter().next()) else {
                    writeln!(output, "[已取消]")?;
                    return Ok(());
                };
                let output_path =
                    unique_workbook_candidate(&input_path, &global.output_dir, "print");
                let action = if choice == "1" {
                    "comment-keep-format"
                } else {
                    "comment-fast-copy"
                };
                let plan_path =
                    unique_menu_output_dir(&global.output_dir, action).join("plan.json");
                let mut arguments = config_cli_arguments(config_path);
                arguments.extend([
                    "--input".to_owned(),
                    input_path.to_string_lossy().into_owned(),
                    "--output".to_owned(),
                    output_path.to_string_lossy().into_owned(),
                    "--drop-comments".to_owned(),
                    (choice == "2").to_string(),
                    "--hide-zero".to_owned(),
                    (choice == "2").to_string(),
                    "--plan-output".to_owned(),
                    plan_path.to_string_lossy().into_owned(),
                ]);
                run_print_comment_plan(arguments)?;
                return Ok(());
            }
            "3" => {
                let output_dir = unique_menu_output_dir(&global.output_dir, "print-config");
                let mut arguments = config_cli_arguments(config_path);
                arguments.extend([
                    "--output-dir".to_owned(),
                    output_dir.to_string_lossy().into_owned(),
                ]);
                run_print_config_plan(arguments)?;
                return Ok(());
            }
            "4" => {
                run_print_config_precheck(config_cli_arguments(config_path))?;
                return Ok(());
            }
            "0" => return Ok(()),
            _ => writeln!(output, "仅接受 1、2、3、4 或 0")?,
        }
    }
}

fn run_interactive_convert_menu<R, W>(
    input: &mut R,
    output: &mut W,
    config_path: &Path,
    global: &GlobalConfig,
    allow_native_dialog: bool,
) -> Result<()>
where
    R: BufRead,
    W: Write,
{
    loop {
        writeln!(output, "------ 转换工具 ------")?;
        writeln!(output, "  1  批量重命名文件")?;
        writeln!(output, "  2  批量 Excel 格式转换（按配置，调用已选引擎）")?;
        writeln!(output, "  3  批量 Word 格式转换（按配置，调用已选引擎）")?;
        writeln!(output, "  4  批量修改 Sheet 名（生成计划）")?;
        writeln!(output, "  5  根据文件内容添加前缀")?;
        writeln!(output, "  6  根据文件内容取消前缀")?;
        writeln!(output, "  7  按局部映射重命名文件")?;
        writeln!(output, "  8  提交缺失检查")?;
        writeln!(output, "  9  添加文件排序前缀")?;
        writeln!(output, " 10  取消文件排序前缀")?;
        writeln!(output, "  0  返回")?;
        writeln!(output, "----------------------")?;
        let Some(choice) = read_menu_choice(input, output, "子菜单选择> ")? else {
            return Ok(());
        };
        let operation = match choice.as_str() {
            "1" => Some(("批量重命名文件", NativeFileFilter::AllFiles)),
            "2" => Some(("批量转换 Excel 文件", NativeFileFilter::ExcelConversion)),
            "3" => Some(("批量转换 Word 文件", NativeFileFilter::WordDocuments)),
            "4" => Some(("批量修改 Sheet 名", NativeFileFilter::XlsxOnly)),
            "5" | "6" => Some(("按工作簿内容重命名文件", NativeFileFilter::XlsxOrXlsm)),
            "7" => Some(("按局部映射重命名文件", NativeFileFilter::AllFiles)),
            "8" => Some(("检查提交文件", NativeFileFilter::AllFiles)),
            "9" => Some(("添加文件排序前缀", NativeFileFilter::AllFiles)),
            "10" => Some(("取消文件排序前缀", NativeFileFilter::AllFiles)),
            "0" => return Ok(()),
            _ => {
                writeln!(output, "仅接受 1 至 10 或 0")?;
                continue;
            }
        };
        let (title, filter) = operation.expect("validated conversion menu choice");
        let Some(paths) =
            choose_source_files_or_folders(title, filter, input, output, allow_native_dialog)?
        else {
            writeln!(output, "[已取消]")?;
            return Ok(());
        };
        if paths.is_empty() {
            writeln!(output, "[提示] 所选范围内没有符合格式的文件。")?;
            return Ok(());
        }

        let mut arguments = config_cli_arguments(config_path);
        match choice.as_str() {
            "1" => {
                arguments.push("--log-dir".to_owned());
                arguments.push(global.log_dir.to_string_lossy().into_owned());
                arguments.extend(
                    paths
                        .into_iter()
                        .map(|path| path.to_string_lossy().into_owned()),
                );
                run_batch_rename(arguments)?;
            }
            "2" => {
                arguments.extend(
                    paths
                        .into_iter()
                        .map(|path| path.to_string_lossy().into_owned()),
                );
                run_convert_batch_cli(arguments, true)?;
            }
            "3" => {
                arguments.extend(
                    paths
                        .into_iter()
                        .map(|path| path.to_string_lossy().into_owned()),
                );
                run_convert_batch_cli(arguments, false)?;
            }
            "4" => {
                let plan_dir = unique_menu_action_dir(&global.output_dir, "sheet-rename");
                fs::create_dir_all(&plan_dir)?;
                let mut planned = 0usize;
                let mut skipped = 0usize;
                for (index, path) in paths.into_iter().enumerate() {
                    if !path
                        .extension()
                        .and_then(|extension| extension.to_str())
                        .is_some_and(|extension| extension.eq_ignore_ascii_case("xlsx"))
                    {
                        writeln!(
                            output,
                            "[跳过] {}：Sheet 改名候选当前只支持 .xlsx",
                            path.display()
                        )?;
                        skipped += 1;
                        continue;
                    }
                    let candidate =
                        unique_workbook_candidate(&path, &global.output_dir, "sheet-rename");
                    let plan_path = plan_dir.join(format!(
                        "{index:04}-{}.plan.json",
                        path.file_stem().unwrap_or_default().to_string_lossy()
                    ));
                    let mut arguments = config_cli_arguments(config_path);
                    arguments.extend([
                        "--input".to_owned(),
                        path.to_string_lossy().into_owned(),
                        "--output".to_owned(),
                        candidate.to_string_lossy().into_owned(),
                        "--plan-output".to_owned(),
                        plan_path.to_string_lossy().into_owned(),
                    ]);
                    run_rename_sheets_plan(arguments)?;
                    planned += 1;
                }
                writeln!(
                    output,
                    "[完成] Sheet 改名计划={planned}，跳过={skipped}；未执行工作簿修改。"
                )?;
            }
            "5" => {
                arguments.extend(
                    paths
                        .into_iter()
                        .map(|path| path.to_string_lossy().into_owned()),
                );
                run_content_prefix_rename_cli(arguments)?;
            }
            "6" => {
                arguments.extend(
                    paths
                        .into_iter()
                        .map(|path| path.to_string_lossy().into_owned()),
                );
                run_content_prefix_unrename_cli(arguments)?;
            }
            "7" => {
                arguments.extend(
                    paths
                        .into_iter()
                        .map(|path| path.to_string_lossy().into_owned()),
                );
                run_local_filename_map(arguments)?;
            }
            "8" => {
                arguments.push("--output-dir".to_owned());
                arguments.push(global.output_dir.to_string_lossy().into_owned());
                arguments.extend(
                    paths
                        .into_iter()
                        .map(|path| path.to_string_lossy().into_owned()),
                );
                run_submission_check_cli(arguments)?;
            }
            "9" | "10" => {
                arguments.push("--output-dir".to_owned());
                arguments.push(global.output_dir.to_string_lossy().into_owned());
                arguments.push("--log-dir".to_owned());
                arguments.push(global.log_dir.to_string_lossy().into_owned());
                arguments.extend(
                    paths
                        .into_iter()
                        .map(|path| path.to_string_lossy().into_owned()),
                );
                if choice == "9" {
                    run_file_sort_prefix(arguments)?;
                } else {
                    run_remove_file_sort_prefix(arguments)?;
                }
            }
            _ => unreachable!("conversion menu operation was validated"),
        }
        return Ok(());
    }
}

fn run_interactive_summary_menu<R, W>(
    input: &mut R,
    output: &mut W,
    global: &GlobalConfig,
    log_view: &str,
    allow_native_dialog: bool,
) -> Result<()>
where
    R: BufRead,
    W: Write,
{
    loop {
        writeln!(output, "------ 汇总工具 ------")?;
        writeln!(output, "  1  按使用区域汇总（全量）")?;
        writeln!(output, "  2  按批注汇总（模板+源文件）")?;
        writeln!(output, "  0  返回")?;
        writeln!(output, "----------------------")?;
        let Some(choice) = read_menu_choice(input, output, "子菜单选择> ")? else {
            return Ok(());
        };
        match choice.as_str() {
            "1" => {
                let Some(sources) = choose_workbooks(
                    "选择按使用区域汇总的源文件（可多选）",
                    true,
                    input,
                    output,
                    allow_native_dialog,
                )?
                else {
                    writeln!(output, "[已取消]")?;
                    return Ok(());
                };
                let mut arguments = summary_cli_arguments(global, log_view);
                arguments.extend(
                    sources
                        .into_iter()
                        .map(|path| path.to_string_lossy().into_owned()),
                );
                run_summary_used_range_cli(arguments)?;
                return Ok(());
            }
            "2" => {
                let Some(template) = choose_workbooks(
                    "选择批注汇总模板文件",
                    false,
                    input,
                    output,
                    allow_native_dialog,
                )?
                .and_then(|paths| paths.into_iter().next()) else {
                    writeln!(output, "[已取消]")?;
                    return Ok(());
                };
                let Some(sources) = choose_workbooks(
                    "选择按批注汇总的源文件（可多选）",
                    true,
                    input,
                    output,
                    allow_native_dialog,
                )?
                else {
                    writeln!(output, "[已取消]")?;
                    return Ok(());
                };
                let mut arguments = summary_cli_arguments(global, log_view);
                arguments.push("--template".to_owned());
                arguments.push(template.to_string_lossy().into_owned());
                arguments.extend(
                    sources
                        .into_iter()
                        .map(|path| path.to_string_lossy().into_owned()),
                );
                run_summary_comments_cli(arguments)?;
                return Ok(());
            }
            "0" => return Ok(()),
            _ => writeln!(output, "输入无效，请重试。")?,
        }
    }
}

fn run_interactive_batch_menu<R, W>(
    input: &mut R,
    output: &mut W,
    config_path: &Path,
    global: &GlobalConfig,
    log_view: &str,
    allow_native_dialog: bool,
    capabilities: Option<&Path>,
    platform: Option<&str>,
) -> Result<()>
where
    R: BufRead,
    W: Write,
{
    loop {
        writeln!(output, "------ 一二批处理工具 ------")?;
        writeln!(output, "  1  拆分村镇银行数据（Rust 计划/能力门控执行）")?;
        writeln!(output, "  2  机构名称规范化（Rust 计划/能力门控执行）")?;
        writeln!(
            output,
            "  3  删除分机构表中的外资行（Rust 计划/能力门控执行）"
        )?;
        writeln!(
            output,
            "  4  修改外汇工作表页眉单位（Rust 计划/能力门控提交）"
        )?;
        writeln!(output, "  5  地区总分校验（读取并生成核对结果）")?;
        writeln!(output, "  0  返回")?;
        writeln!(output, "----------------------------")?;
        let Some(choice) = read_menu_choice(input, output, "子菜单选择> ")? else {
            return Ok(());
        };
        match choice.as_str() {
            "1" => {
                let Some(current) = choose_single_xlsx_workbook(
                    "选择本期村镇银行数据工作簿",
                    input,
                    output,
                    allow_native_dialog,
                )?
                else {
                    writeln!(output, "[已取消] 未选择本期工作簿。")?;
                    return Ok(());
                };
                let Some(previous) = choose_single_xlsx_workbook(
                    "选择去年同期村镇银行数据工作簿",
                    input,
                    output,
                    allow_native_dialog,
                )?
                else {
                    writeln!(output, "[已取消] 未选择上期工作簿。")?;
                    return Ok(());
                };
                let Some(selected) = choose_source_files_or_folders(
                    "选择需要拆分村镇银行数据的目标工作簿",
                    NativeFileFilter::XlsxOnly,
                    input,
                    output,
                    allow_native_dialog,
                )?
                else {
                    writeln!(output, "[已取消] 未选择目标工作簿。")?;
                    return Ok(());
                };
                let mut targets = Vec::new();
                for path in selected {
                    if path
                        .extension()
                        .is_some_and(|extension| extension.eq_ignore_ascii_case("xlsx"))
                    {
                        targets.push(path);
                    } else {
                        writeln!(
                            output,
                            "[跳过] {}：PT-7-1 当前只接受 .xlsx；未验收的格式请勿直接修改。",
                            path.display()
                        )?;
                    }
                }
                if targets.is_empty() {
                    writeln!(output, "[提示] 没有可处理的 .xlsx 目标工作簿。")?;
                    return Ok(());
                }
                if capabilities.is_none() {
                    writeln!(
                        output,
                        "[计划模式] 当前未提供已验收的 ExcelModifier 能力文件；原目标不会修改。"
                    )?;
                } else {
                    writeln!(
                        output,
                        "[执行模式] 将按能力文件生成并验证候选，再安全替换目标并保留备份。"
                    )?;
                }
                let mut arguments = vec![
                    "--current".to_owned(),
                    current.to_string_lossy().into_owned(),
                    "--previous".to_owned(),
                    previous.to_string_lossy().into_owned(),
                    "--output-dir".to_owned(),
                    global.output_dir.to_string_lossy().into_owned(),
                    "--log-dir".to_owned(),
                    global.log_dir.to_string_lossy().into_owned(),
                    "--log-view".to_owned(),
                    log_view.to_owned(),
                ];
                if let (Some(capabilities), Some(platform)) = (capabilities, platform) {
                    arguments.extend([
                        "--capabilities".to_owned(),
                        capabilities.to_string_lossy().into_owned(),
                        "--platform".to_owned(),
                        platform.to_owned(),
                        "--commit-target".to_owned(),
                    ]);
                }
                arguments.extend(
                    targets
                        .into_iter()
                        .map(|path| path.to_string_lossy().into_owned()),
                );
                run_village_split_batch(arguments)?;
                return Ok(());
            }
            "2" => {
                let Some(selected) = choose_source_files_or_folders(
                    "选择需要规范化机构名称的目标工作簿",
                    NativeFileFilter::XlsxOnly,
                    input,
                    output,
                    allow_native_dialog,
                )?
                else {
                    writeln!(output, "[已取消]")?;
                    return Ok(());
                };
                let mut targets = Vec::new();
                for path in selected {
                    if path
                        .extension()
                        .is_some_and(|extension| extension.eq_ignore_ascii_case("xlsx"))
                    {
                        targets.push(path);
                    } else {
                        writeln!(
                            output,
                            "[跳过] {}：PT-7-2 当前只接受已验证的 .xlsx；未验收格式不执行修改。",
                            path.display()
                        )?;
                    }
                }
                if targets.is_empty() {
                    writeln!(output, "[提示] 没有可处理的 .xlsx 目标工作簿。")?;
                    return Ok(());
                }
                if capabilities.is_none() {
                    writeln!(output, "[计划模式] 从配置读取机构映射；原目标不会修改。")?;
                } else {
                    writeln!(
                        output,
                        "[执行模式] 将按 capability 生成并验证候选，再安全替换目标并保留备份。"
                    )?;
                }

                let plan_dir = global.output_dir.join("plans").join("PT-7-2");
                let mut arguments = vec![
                    "--config".to_owned(),
                    config_path.to_string_lossy().into_owned(),
                    "--view".to_owned(),
                    log_view.to_owned(),
                ];
                if log_view != "off" {
                    let nonce = SystemTime::now()
                        .duration_since(UNIX_EPOCH)
                        .map(|duration| duration.as_nanos())
                        .unwrap_or_default();
                    let report_path = global.log_dir.join(format!(
                        "PT-7-2_{}_{}_{}.json",
                        Local::now().format("%Y%m%dT%H%M%S"),
                        std::process::id(),
                        nonce
                    ));
                    arguments.extend([
                        "--report-output".to_owned(),
                        report_path.to_string_lossy().into_owned(),
                    ]);
                }
                if let (Some(capabilities), Some(platform)) = (capabilities, platform) {
                    arguments.extend([
                        "--apply".to_owned(),
                        "--capabilities".to_owned(),
                        capabilities.to_string_lossy().into_owned(),
                        "--platform".to_owned(),
                        platform.to_owned(),
                    ]);
                } else {
                    arguments.extend([
                        "--plan-dir".to_owned(),
                        plan_dir.to_string_lossy().into_owned(),
                    ]);
                }
                arguments.extend(
                    targets
                        .into_iter()
                        .map(|path| path.to_string_lossy().into_owned()),
                );
                run_institution_normalize_batch(arguments)?;
                return Ok(());
            }
            "3" => {
                let foreign = pytools_core::config::load_foreign_banks(config_path)?;
                if foreign.is_empty() {
                    writeln!(
                        output,
                        "[已中止] 外资行配置为空，请在机构映射表 C 列将外资行标记为 1。"
                    )?;
                    return Ok(());
                }
                let Some(selected) = choose_source_files_or_folders(
                    "选择要生成外发用副本的工作簿（删除分机构表的外资行）",
                    NativeFileFilter::XlsxOnly,
                    input,
                    output,
                    allow_native_dialog,
                )?
                else {
                    writeln!(output, "[已取消]")?;
                    return Ok(());
                };
                let mut targets = Vec::new();
                for path in selected {
                    if path
                        .extension()
                        .is_some_and(|extension| extension.eq_ignore_ascii_case("xlsx"))
                    {
                        targets.push(path);
                    } else {
                        writeln!(
                            output,
                            "[跳过] {}：PT-7-3 当前只接受 .xlsx；未验收格式请先显式转换。",
                            path.display()
                        )?;
                    }
                }
                if targets.is_empty() {
                    writeln!(output, "[提示] 没有可处理的 .xlsx 工作簿。")?;
                    return Ok(());
                }
                if capabilities.is_none() {
                    writeln!(output, "[计划模式] 默认只生成修改计划；原工作簿不会修改。")?;
                } else {
                    writeln!(
                        output,
                        "[执行模式] 将通过已提供的 ExcelModifier capability 生成外发副本；原工作簿不会修改。"
                    )?;
                }

                let mut arguments = vec![
                    "--config".to_owned(),
                    config_path.to_string_lossy().into_owned(),
                    "--view".to_owned(),
                    log_view.to_owned(),
                    "--log-dir".to_owned(),
                    global.log_dir.to_string_lossy().into_owned(),
                ];
                if log_view != "off" {
                    let nonce = SystemTime::now()
                        .duration_since(UNIX_EPOCH)
                        .map(|duration| duration.as_nanos())
                        .unwrap_or_default();
                    let report_path = global.log_dir.join(format!(
                        "PT-7-3_{}_{}_{}.json",
                        Local::now().format("%Y%m%dT%H%M%S"),
                        std::process::id(),
                        nonce
                    ));
                    arguments.extend([
                        "--report-output".to_owned(),
                        report_path.to_string_lossy().into_owned(),
                    ]);
                }
                if let (Some(capabilities), Some(platform)) = (capabilities, platform) {
                    arguments.extend([
                        "--capabilities".to_owned(),
                        capabilities.to_string_lossy().into_owned(),
                        "--platform".to_owned(),
                        platform.to_owned(),
                    ]);
                } else {
                    arguments.extend([
                        "--plan-dir".to_owned(),
                        global
                            .output_dir
                            .join("plans")
                            .join("PT-7-3")
                            .to_string_lossy()
                            .into_owned(),
                    ]);
                }
                arguments.extend(
                    targets
                        .into_iter()
                        .map(|path| path.to_string_lossy().into_owned()),
                );
                run_remove_foreign_bank_batch(arguments)?;
                return Ok(());
            }
            "4" => {
                let Some(selected) = choose_workbooks(
                    "选择要修改外汇页眉的 Excel 文件（可多选）",
                    true,
                    input,
                    output,
                    allow_native_dialog,
                )?
                else {
                    writeln!(output, "[已取消] 未选择外汇工作簿。")?;
                    return Ok(());
                };
                let mut targets = Vec::new();
                for path in selected {
                    if path
                        .extension()
                        .is_some_and(|extension| extension.eq_ignore_ascii_case("xlsx"))
                    {
                        targets.push(path);
                    } else {
                        writeln!(
                            output,
                            "[跳过] {}：PT-7-4 当前仅验收 .xlsx；.xls 请先显式转换，含 VBA 的 .xlsm 暂不修改。",
                            path.display()
                        )?;
                    }
                }
                if targets.is_empty() {
                    writeln!(output, "[提示] 没有可处理的 .xlsx 工作簿。")?;
                    return Ok(());
                }
                if capabilities.is_none() {
                    writeln!(
                        output,
                        "[计划模式] 默认只保存页眉修改计划；原工作簿不会修改。"
                    )?;
                } else {
                    writeln!(
                        output,
                        "[执行模式] 将通过匹配的 ExcelModifier 验证候选后安全替换原工作簿；失败时使用临时备份回滚。"
                    )?;
                }
                let mut arguments = vec![
                    "--output-dir".to_owned(),
                    global.output_dir.to_string_lossy().into_owned(),
                    "--log-dir".to_owned(),
                    global.log_dir.to_string_lossy().into_owned(),
                    "--view".to_owned(),
                    log_view.to_owned(),
                ];
                if let (Some(capabilities), Some(platform)) = (capabilities, platform) {
                    arguments.extend([
                        "--commit-target".to_owned(),
                        "--capabilities".to_owned(),
                        capabilities.to_string_lossy().into_owned(),
                        "--platform".to_owned(),
                        platform.to_owned(),
                    ]);
                }
                arguments.extend(
                    targets
                        .into_iter()
                        .map(|path| path.to_string_lossy().into_owned()),
                );
                run_fx_header_batch(arguments)?;
                return Ok(());
            }
            "5" => {
                let Some(selected) = choose_source_files_or_folders(
                    "选择要做地区总分校验的 Excel 文件",
                    NativeFileFilter::XlsxOrXlsm,
                    input,
                    output,
                    allow_native_dialog,
                )?
                else {
                    writeln!(output, "[已取消]")?;
                    return Ok(());
                };
                let mut sources = Vec::new();
                for path in selected {
                    if path.extension().is_some_and(|extension| {
                        extension.eq_ignore_ascii_case("xlsx")
                            || extension.eq_ignore_ascii_case("xlsm")
                    }) {
                        sources.push(path);
                    } else {
                        writeln!(
                            output,
                            "[跳过] {}：当前 Rust 核对入口只支持 .xlsx/.xlsm；请先显式转换为 .xlsx",
                            path.display()
                        )?;
                    }
                }
                if sources.is_empty() {
                    writeln!(output, "[提示] 没有可处理的 .xlsx/.xlsm 文件。")?;
                    return Ok(());
                }
                let mut arguments = config_cli_arguments(config_path);
                arguments.extend([
                    "--output-dir".to_owned(),
                    global.output_dir.to_string_lossy().into_owned(),
                    "--log-dir".to_owned(),
                    global.log_dir.to_string_lossy().into_owned(),
                ]);
                arguments.extend(
                    sources
                        .into_iter()
                        .map(|path| path.to_string_lossy().into_owned()),
                );
                run_region_sum(arguments)?;
                return Ok(());
            }
            "0" => return Ok(()),
            _ => writeln!(output, "仅接受 1、2、3、4、5 或 0")?,
        }
    }
}

fn summary_cli_arguments(global: &GlobalConfig, log_view: &str) -> Vec<String> {
    vec![
        "--output-dir".to_owned(),
        global.output_dir.to_string_lossy().into_owned(),
        "--log-dir".to_owned(),
        global.log_dir.to_string_lossy().into_owned(),
        "--log-view".to_owned(),
        log_view.to_owned(),
    ]
}

fn read_menu_choice<R, W>(input: &mut R, output: &mut W, prompt: &str) -> Result<Option<String>>
where
    R: BufRead,
    W: Write,
{
    write!(output, "{prompt}")?;
    output.flush()?;
    let mut line = String::new();
    if input.read_line(&mut line)? == 0 {
        return Ok(None);
    }
    Ok(Some(line.trim().to_owned()))
}

fn choose_workbooks<R, W>(
    title: &str,
    allow_multiple: bool,
    input: &mut R,
    output: &mut W,
    allow_native_dialog: bool,
) -> Result<Option<Vec<PathBuf>>>
where
    R: BufRead,
    W: Write,
{
    choose_input_paths(
        title,
        allow_multiple,
        NativeFileFilter::ExcelWorkbooks,
        input,
        output,
        allow_native_dialog,
    )
}

fn choose_single_xlsx_workbook<R, W>(
    title: &str,
    input: &mut R,
    output: &mut W,
    allow_native_dialog: bool,
) -> Result<Option<PathBuf>>
where
    R: BufRead,
    W: Write,
{
    let Some(mut selected) = choose_input_paths(
        title,
        false,
        NativeFileFilter::XlsxOnly,
        input,
        output,
        allow_native_dialog,
    )?
    else {
        return Ok(None);
    };
    if selected.len() != 1 {
        writeln!(output, "[提示] 此步骤需要且只需要选择一个 .xlsx 工作簿。")?;
        return Ok(None);
    }
    let path = selected.remove(0);
    if !path
        .extension()
        .is_some_and(|extension| extension.eq_ignore_ascii_case("xlsx"))
    {
        writeln!(
            output,
            "[跳过] {}：PT-7-1 当前只接受 .xlsx；未验收的格式请勿直接修改。",
            path.display()
        )?;
        return Ok(None);
    }
    Ok(Some(path))
}

#[derive(Clone, Copy)]
enum NativeFileFilter {
    ExcelWorkbooks,
    ExcelConversion,
    XlsxOrXlsm,
    XlsxOnly,
    WordDocuments,
    AllFiles,
}

impl NativeFileFilter {
    fn dialog_filter(self) -> &'static str {
        match self {
            Self::ExcelWorkbooks => {
                "Excel files (*.xlsx;*.xlsm;*.xls)\0*.xlsx;*.xlsm;*.xls\0All files (*.*)\0*.*\0\0"
            }
            Self::ExcelConversion => {
                "Excel files (*.csv;*.et;*.xls;*.xlsx;*.xlsm;*.xlsb)\0*.csv;*.et;*.xls;*.xlsx;*.xlsm;*.xlsb\0All files (*.*)\0*.*\0\0"
            }
            Self::XlsxOrXlsm => {
                "Excel files (*.xlsx;*.xlsm)\0*.xlsx;*.xlsm\0All files (*.*)\0*.*\0\0"
            }
            Self::XlsxOnly => "Excel files (*.xlsx)\0*.xlsx\0All files (*.*)\0*.*\0\0",
            Self::WordDocuments => {
                "Word files (*.doc;*.docx)\0*.doc;*.docx\0All files (*.*)\0*.*\0\0"
            }
            Self::AllFiles => "All files (*.*)\0*.*\0\0",
        }
    }

    fn accepts_folder_file(self, path: &Path) -> bool {
        let extension = path
            .extension()
            .and_then(|value| value.to_str())
            .unwrap_or_default()
            .to_ascii_lowercase();
        match self {
            Self::ExcelWorkbooks => matches!(extension.as_str(), "xlsx" | "xlsm" | "xls"),
            Self::ExcelConversion => {
                matches!(
                    extension.as_str(),
                    "csv" | "et" | "xls" | "xlsx" | "xlsm" | "xlsb"
                )
            }
            Self::XlsxOrXlsm => matches!(extension.as_str(), "xlsx" | "xlsm"),
            Self::XlsxOnly => extension == "xlsx",
            Self::WordDocuments => matches!(extension.as_str(), "doc" | "docx"),
            Self::AllFiles => true,
        }
    }
}

fn choose_input_paths<R, W>(
    title: &str,
    allow_multiple: bool,
    filter: NativeFileFilter,
    input: &mut R,
    output: &mut W,
    allow_native_dialog: bool,
) -> Result<Option<Vec<PathBuf>>>
where
    R: BufRead,
    W: Write,
{
    writeln!(output, "→ 正在打开文件选择窗口：{title}")?;
    if allow_native_dialog {
        if let Some(paths) = native_file_dialog(title, allow_multiple, filter) {
            if !paths.is_empty() {
                print_selected_paths(output, &paths)?;
                return Ok(Some(paths));
            }
        }
        writeln!(output, "[提示] 文件弹窗不可用或已取消。")?;
    }
    writeln!(output, "→ 可手动输入路径（分号分隔），或输入 0 取消：")?;
    let Some(raw) = read_menu_choice(input, output, "文件路径> ")? else {
        return Ok(None);
    };
    let Some(paths) = parse_manual_path_list(&raw) else {
        return Ok(None);
    };
    print_selected_paths(output, &paths)?;
    Ok(Some(paths))
}

fn choose_source_files_or_folders<R, W>(
    title: &str,
    filter: NativeFileFilter,
    input: &mut R,
    output: &mut W,
    allow_native_dialog: bool,
) -> Result<Option<Vec<PathBuf>>>
where
    R: BufRead,
    W: Write,
{
    writeln!(output, "→ {title}")?;
    writeln!(output, "  1  选择文件（可多选）")?;
    writeln!(output, "  2  选择文件夹（仅当前层）")?;
    writeln!(output, "  3  选择文件夹（递归子目录）")?;
    writeln!(output, "  0  取消")?;
    let Some(mode) = read_menu_choice(input, output, "选择方式> ")? else {
        return Ok(None);
    };
    if mode == "0" {
        return Ok(None);
    }
    if mode == "1" {
        return choose_input_paths(title, true, filter, input, output, allow_native_dialog);
    }
    anyhow::ensure!(matches!(mode.as_str(), "2" | "3"), "仅接受 1、2、3 或 0");

    let native_folder = allow_native_dialog
        .then(|| native_folder_dialog(title))
        .flatten();
    let mut folders = native_folder.into_iter().collect::<Vec<_>>();
    if folders.is_empty() {
        writeln!(
            output,
            "→ 可手动输入文件夹路径（多个路径用分号分隔），或输入 0 取消："
        )?;
        let Some(raw) = read_menu_choice(input, output, "文件夹路径> ")? else {
            return Ok(None);
        };
        let Some(paths) = parse_manual_path_list(&raw) else {
            return Ok(None);
        };
        folders.extend(paths);
    } else {
        print_selected_paths(output, &folders)?;
        writeln!(output, "→ 可补充其它文件夹路径（分号分隔，留空继续）：")?;
        if let Some(raw) = read_menu_choice(input, output, "补充路径> ")? {
            if let Some(paths) = parse_manual_path_list(&raw) {
                folders.extend(paths);
            }
        }
    }
    let paths = collect_menu_source_files(&folders, mode == "3", filter)?;
    writeln!(
        output,
        "→ 已选 {} 个文件夹，{}共识别文件 {} 个",
        folders.len(),
        if mode == "3" { "递归" } else { "当前层" },
        paths.len()
    )?;
    print_selected_paths(output, &paths)?;
    Ok(Some(paths))
}

fn collect_menu_source_files(
    folders: &[PathBuf],
    recursive: bool,
    filter: NativeFileFilter,
) -> Result<Vec<PathBuf>> {
    let mut files = Vec::new();
    for folder in folders {
        anyhow::ensure!(
            folder.is_dir(),
            "selected folder does not exist or is not a directory: {}",
            folder.display()
        );
        collect_menu_source_files_from_folder(folder, folder, recursive, filter, &mut files)?;
    }
    let mut seen = std::collections::BTreeSet::new();
    files.retain(|path| {
        let identity = fs::canonicalize(path)
            .unwrap_or_else(|_| path.to_path_buf())
            .to_string_lossy()
            .to_lowercase();
        seen.insert(identity)
    });
    Ok(files)
}

fn collect_menu_source_files_from_folder(
    root: &Path,
    folder: &Path,
    recursive: bool,
    filter: NativeFileFilter,
    files: &mut Vec<PathBuf>,
) -> Result<()> {
    let mut entries = fs::read_dir(folder)
        .with_context(|| format!("list selected folder {}", folder.display()))?
        .map(|entry| entry.map(|entry| entry.path()))
        .collect::<std::io::Result<Vec<_>>>()?;
    entries.sort_by_key(|path| {
        path.strip_prefix(root)
            .unwrap_or(path)
            .to_string_lossy()
            .to_lowercase()
    });
    for path in entries {
        let metadata = fs::symlink_metadata(&path)
            .with_context(|| format!("inspect selected path {}", path.display()))?;
        if metadata.file_type().is_symlink() {
            continue;
        }
        if metadata.is_file() {
            if !path
                .file_name()
                .is_some_and(|name| name.to_string_lossy().starts_with("~$"))
                && filter.accepts_folder_file(&path)
            {
                files.push(path);
            }
        } else if recursive && metadata.is_dir() {
            collect_menu_source_files_from_folder(root, &path, true, filter, files)?;
        }
    }
    Ok(())
}

fn parse_manual_path_list(raw: &str) -> Option<Vec<PathBuf>> {
    let raw = raw.trim();
    if raw.is_empty() || raw == "0" {
        return None;
    }
    let paths = raw
        .split(';')
        .map(|part| {
            part.trim()
                .trim_matches(|character| matches!(character, '\"' | '\''))
        })
        .filter(|part| !part.is_empty())
        .map(PathBuf::from)
        .collect::<Vec<_>>();
    (!paths.is_empty()).then_some(paths)
}

fn print_selected_paths<W: Write>(output: &mut W, paths: &[PathBuf]) -> Result<()> {
    writeln!(output, "→ 已选择文件 {} 个：", paths.len())?;
    for (index, path) in paths.iter().enumerate() {
        writeln!(output, "  {}. {}", index + 1, path.display())?;
    }
    Ok(())
}

#[cfg(windows)]
fn native_file_dialog(
    title: &str,
    allow_multiple: bool,
    file_filter: NativeFileFilter,
) -> Option<Vec<PathBuf>> {
    use std::os::windows::ffi::OsStringExt;
    use windows_sys::Win32::UI::Controls::Dialogs::{
        GetOpenFileNameW, OFN_ALLOWMULTISELECT, OFN_EXPLORER, OFN_FILEMUSTEXIST, OFN_HIDEREADONLY,
        OFN_NOCHANGEDIR, OFN_PATHMUSTEXIST, OPENFILENAMEW,
    };

    let mut filter = file_filter
        .dialog_filter()
        .encode_utf16()
        .collect::<Vec<_>>();
    let mut title = title
        .encode_utf16()
        .chain(std::iter::once(0))
        .collect::<Vec<_>>();
    let mut buffer = vec![0u16; 32_768];
    let mut dialog = OPENFILENAMEW::default();
    dialog.lStructSize = std::mem::size_of::<OPENFILENAMEW>() as u32;
    dialog.lpstrFilter = filter.as_mut_ptr();
    dialog.nFilterIndex = 1;
    dialog.lpstrFile = buffer.as_mut_ptr();
    dialog.nMaxFile = buffer.len() as u32;
    dialog.lpstrTitle = title.as_mut_ptr();
    dialog.Flags = OFN_EXPLORER
        | OFN_FILEMUSTEXIST
        | OFN_PATHMUSTEXIST
        | OFN_NOCHANGEDIR
        | OFN_HIDEREADONLY
        | if allow_multiple {
            OFN_ALLOWMULTISELECT
        } else {
            0
        };

    let selected = unsafe { GetOpenFileNameW(&mut dialog) };
    if selected == 0 {
        return None;
    }
    let parts = buffer
        .split(|unit| *unit == 0)
        .take_while(|part| !part.is_empty())
        .collect::<Vec<_>>();
    if parts.is_empty() {
        return None;
    }
    if parts.len() == 1 {
        return Some(vec![PathBuf::from(std::ffi::OsString::from_wide(parts[0]))]);
    }
    let directory = PathBuf::from(std::ffi::OsString::from_wide(parts[0]));
    Some(
        parts[1..]
            .iter()
            .map(|name| directory.join(std::ffi::OsString::from_wide(name)))
            .collect(),
    )
}

#[cfg(not(windows))]
fn native_file_dialog(
    _title: &str,
    _allow_multiple: bool,
    _file_filter: NativeFileFilter,
) -> Option<Vec<PathBuf>> {
    None
}

#[cfg(windows)]
fn native_folder_dialog(title: &str) -> Option<PathBuf> {
    use std::os::windows::ffi::OsStringExt;
    use windows_sys::Win32::System::Com::CoTaskMemFree;
    use windows_sys::Win32::System::Ole::{OleInitialize, OleUninitialize};
    use windows_sys::Win32::UI::Shell::{
        SHBrowseForFolderW, SHGetPathFromIDListW, BIF_EDITBOX, BIF_NEWDIALOGSTYLE,
        BIF_RETURNONLYFSDIRS, BROWSEINFOW,
    };

    let mut display_name = vec![0u16; 260];
    let title = title
        .encode_utf16()
        .chain(std::iter::once(0))
        .collect::<Vec<_>>();
    let mut dialog = BROWSEINFOW::default();
    dialog.pszDisplayName = display_name.as_mut_ptr();
    dialog.lpszTitle = title.as_ptr();
    dialog.ulFlags = BIF_RETURNONLYFSDIRS | BIF_NEWDIALOGSTYLE | BIF_EDITBOX;

    // SHBrowseForFolderW's new-style dialog requires COM on an STA thread.
    // S_FALSE also means initialization succeeded and must be balanced.
    if unsafe { OleInitialize(std::ptr::null()) } < 0 {
        return None;
    }
    let item = unsafe { SHBrowseForFolderW(&dialog) };
    let selected = if item.is_null() {
        None
    } else {
        let mut path = vec![0u16; 32_768];
        let resolved = unsafe { SHGetPathFromIDListW(item, path.as_mut_ptr()) } != 0;
        unsafe { CoTaskMemFree(item.cast()) };
        if !resolved {
            None
        } else {
            let length = path
                .iter()
                .position(|unit| *unit == 0)
                .unwrap_or(path.len());
            Some(PathBuf::from(std::ffi::OsString::from_wide(
                &path[..length],
            )))
        }
    };
    unsafe { OleUninitialize() };
    selected
}

#[cfg(not(windows))]
fn native_folder_dialog(_title: &str) -> Option<PathBuf> {
    None
}

fn run_quick_cli(arguments: Vec<String>, action: &str) -> Result<()> {
    let options = parse_cli_arguments(arguments)?;
    let config = required_path(&options, "config")?;
    let global = load_global(config)?;
    match action {
        "config" => open_config(
            options
                .values
                .get("config")
                .map(Path::new)
                .context("--config is required")?,
        )?,
        "latest" => {
            let path = open_latest_output(&global.output_dir)?;
            println!("{}", serde_json::json!({"latest":path}));
        }
        "output" => open_output_directory(&global.output_dir)?,
        _ => unreachable!(),
    };
    Ok(())
}

fn run_config_init_cli(arguments: Vec<String>) -> Result<()> {
    let options = parse_cli_arguments(arguments)?;
    let path = required_path(&options, "path")?;
    let result = initialize_new_config(&path)?;
    println!("{}", serde_json::to_string_pretty(&result)?);
    Ok(())
}
fn run_config_view_cli(arguments: Vec<String>) -> Result<()> {
    let options = parse_cli_arguments(arguments)?;
    let path = required_path(&options, "config")?;
    let result = view_config(path)?;
    println!("{}", serde_json::to_string_pretty(&result)?);
    Ok(())
}

fn run_config_repair_plan_cli(arguments: Vec<String>) -> Result<()> {
    let options = parse_cli_arguments(arguments)?;
    let input = required_path(&options, "input")?;
    let output = required_path(&options, "output")?;
    let capabilities = options.values.get("capabilities");
    let platform = options.values.get("platform");
    anyhow::ensure!(
        capabilities.is_some() == platform.is_some(),
        "--capabilities and --platform must be provided together"
    );
    let Some(plan) = build_config_repair_plan(&input, &output)? else {
        println!(
            "{{\"status\":\"already_complete\",\"input\":{}}}",
            serde_json::to_string(&input.to_string_lossy())?
        );
        return Ok(());
    };
    if let Some(plan_path) = options.values.get("plan-output") {
        fs::write(plan_path, serde_json::to_vec_pretty(&plan)?)?;
    }
    if let (Some(capabilities), Some(platform)) = (capabilities, platform) {
        let execution = apply_plan_with_capability_file(&plan, capabilities, platform)?;
        println!(
            "{}",
            serde_json::to_string_pretty(&serde_json::json!({
                "plan": plan,
                "plan_file": options.values.get("plan-output"),
                "execution": execution,
            }))?
        );
    } else {
        println!("{}", serde_json::to_string_pretty(&plan)?);
    }
    Ok(())
}

fn run_dedup_comment_plan(arguments: Vec<String>) -> Result<()> {
    let options = parse_cli_arguments(arguments)?;
    let input = required_path(&options, "input")?;
    let output = required_path(&options, "output")?;
    let action = parse_duplicate_action(
        options
            .values
            .get("action")
            .map(String::as_str)
            .unwrap_or("check"),
    )?;
    let report = build_comment_duplicate_plan(&input, &output, action)?;
    let capabilities = options.values.get("capabilities");
    let platform = options.values.get("platform");
    anyhow::ensure!(
        capabilities.is_some() == platform.is_some(),
        "--capabilities and --platform must be provided together"
    );
    let mut result = serde_json::json!({
        "workbook": report.workbook,
        "hit_sheets": report.hit_sheets,
        "duplicate_rows": report.duplicate_rows,
        "skipped_tasks": report.skipped_tasks,
        "plan": report.plan,
        "execution": if capabilities.is_some() { "not_needed" } else { "plan_only" },
    });
    let planned = result.get("plan").filter(|plan| !plan.is_null()).cloned();
    if let Some(plan_value) = planned {
        if let Some(plan_path) = options.values.get("plan-output") {
            if let Some(parent) = Path::new(plan_path)
                .parent()
                .filter(|path| !path.as_os_str().is_empty())
            {
                fs::create_dir_all(parent)?;
            }
            fs::write(plan_path, serde_json::to_vec_pretty(&plan_value)?)?;
            result["plan_file"] = serde_json::json!(plan_path);
        }
        if let (Some(capabilities), Some(platform)) = (capabilities, platform) {
            let plan: excel_editor::ExcelEditPlan = serde_json::from_value(plan_value)?;
            let execution = apply_plan_with_capability_file(&plan, capabilities, platform)?;
            result["execution"] = serde_json::json!(execution);
        }
    }
    println!("{}", serde_json::to_string_pretty(&result)?);
    Ok(())
}

fn run_dedup_config_plan(arguments: Vec<String>) -> Result<()> {
    let options = parse_cli_arguments(arguments)?;
    let config = required_path(&options, "config")?;
    let output_dir = options
        .values
        .get("output-dir")
        .map(PathBuf::from)
        .unwrap_or(load_global(&config)?.output_dir);
    let action = parse_duplicate_action(
        options
            .values
            .get("action")
            .map(String::as_str)
            .unwrap_or("check"),
    )?;
    let tasks = pytools_core::config::load_dedup_tasks(&config)?;
    let mut outputs = std::collections::BTreeMap::new();
    for task in tasks.iter().filter(|task| task.enabled) {
        if let Some(input) = &task.source_wb {
            outputs
                .entry(input.clone())
                .or_insert_with(|| unique_workbook_candidate(input, &output_dir, "dedup"));
        }
    }
    fs::create_dir_all(&output_dir)?;
    let reports = build_config_duplicate_plan(&tasks, &outputs, action)?;
    let capabilities = options.values.get("capabilities");
    let platform = options.values.get("platform");
    anyhow::ensure!(
        capabilities.is_some() == platform.is_some(),
        "--capabilities and --platform must be provided together"
    );
    let mut results = Vec::with_capacity(reports.len());
    for (index, report) in reports.iter().enumerate() {
        let mut plan_path = None;
        let mut execution = if capabilities.is_some() {
            serde_json::json!("not_needed")
        } else {
            serde_json::json!("plan_only")
        };
        if let Some(plan) = &report.plan {
            let path = output_dir.join(format!(
                "dedup-plan-{}-{index}.json",
                chrono::Local::now().format("%Y%m%d_%H%M%S")
            ));
            fs::write(&path, serde_json::to_vec_pretty(plan)?)?;
            plan_path = Some(path.clone());
            if let (Some(capabilities), Some(platform)) = (capabilities, platform) {
                execution = serde_json::json!(apply_plan_with_capability_file(
                    plan,
                    capabilities,
                    platform
                )?);
            }
        }
        results.push(serde_json::json!({
            "report": report,
            "plan_file": plan_path,
            "execution": execution,
        }));
    }
    println!("{}", serde_json::to_string_pretty(&results)?);
    Ok(())
}

fn run_dedup_precheck(arguments: Vec<String>) -> Result<()> {
    let options = parse_cli_arguments(arguments)?;
    let config = required_path(&options, "config")?;
    let tasks = pytools_core::config::load_dedup_tasks(config)?;
    let report = precheck_dedup_tasks(&tasks);
    println!("{}", serde_json::to_string_pretty(&report)?);
    if report.fail > 0 {
        bail!("dedup precheck found {} failure(s)", report.fail);
    }
    Ok(())
}

fn run_dedup_append_plan(arguments: Vec<String>) -> Result<()> {
    let options = parse_cli_arguments(arguments)?;
    let config = required_path(&options, "config")?;
    let global = load_global(&config)?;
    let output_dir = options
        .values
        .get("output-dir")
        .map(PathBuf::from)
        .unwrap_or(global.output_dir);
    let tasks = pytools_core::config::load_dedup_tasks(&config)?;
    let reports = build_append_by_config_plans(&tasks, &output_dir)?;
    fs::create_dir_all(&output_dir)?;
    let timestamp = Local::now().format("%Y%m%d_%H%M%S");
    let mut output = Vec::new();
    for (index, report) in reports.iter().enumerate() {
        let plan_path = output_dir.join(format!("dedup-append-{timestamp}-{index}.json"));
        fs::write(&plan_path, serde_json::to_vec_pretty(report)?)?;
        output.push(serde_json::json!({
            "target": report.target,
            "candidate": report.candidate,
            "target_existed": report.target_existed,
            "task_ok": report.task_ok,
            "task_skipped": report.task_skipped,
            "appended_rows": report.appended_rows,
            "duplicate_rows_filtered": report.duplicate_rows_filtered,
            "merged_tasks": report.merged_tasks,
            "issues": report.issues,
            "plan_file": plan_path,
            "execution": if report.target_existed { "requires apply-edit-plan capability evidence" } else { "run dedup-append-execute" },
        }));
    }
    println!("{}", serde_json::to_string_pretty(&output)?);
    Ok(())
}

fn run_dedup_append_execute(arguments: Vec<String>) -> Result<()> {
    let run_started = Instant::now();
    let (commit_target, arguments) = extract_boolean_flag(arguments, "commit-target")?;
    let options = parse_cli_arguments(arguments)?;
    let plan_path = required_path(&options, "plan")?;
    let plan: AppendTargetPlan = serde_json::from_slice(&fs::read(&plan_path)?)
        .with_context(|| format!("decode append plan {}", plan_path.display()))?;
    if plan.target_existed {
        let edit_plan = plan
            .edit_plan
            .as_ref()
            .context("append plan does not contain an edit plan")?;
        edit_plan.validate().context("stage=edit_plan_validation")?;
        let capabilities = required_path(&options, "capabilities")?;
        let platform = options
            .values
            .get("platform")
            .context("--platform is required")?;
        let expected_target_hash = if commit_target {
            let expected = plan.target_sha256.as_deref().context(
                "append plan has no target SHA-256; regenerate the plan before committing",
            )?;
            let actual = sha256_file(&plan.target)?;
            anyhow::ensure!(
                actual == expected,
                "configured target changed after planning; refusing to execute the commit: {}",
                plan.target.display()
            );
            Some(expected.to_owned())
        } else {
            None
        };
        let staging = if plan.replacement_workbook.is_some() {
            Some(TemporaryFile::new(prepare_wide_target_replacement_source(
                &plan,
            )?))
        } else {
            None
        };
        let candidate_preexisted = plan.candidate.exists();
        let edit_started = Instant::now();
        let execution = match apply_plan_with_capability_file(edit_plan, capabilities, platform) {
            Ok(execution) => execution,
            Err(error) => {
                if !candidate_preexisted && plan.candidate.exists() {
                    let _ = fs::remove_file(&plan.candidate);
                }
                return Err(error).context("stage=excel_modifier");
            }
        };
        let edit_elapsed_ms = edit_started.elapsed().as_millis();
        let mut stages = vec![serde_json::json!({
            "name": "excel_modify",
            "status": "succeeded",
            "engine": execution.engine,
            "elapsed_ms": edit_elapsed_ms
        })];
        if let Some(staging) = &staging {
            let validation_started = Instant::now();
            if let Err(error) = validate_wide_target_replacement_candidate(&plan, staging.path()) {
                if !candidate_preexisted && plan.candidate.exists() {
                    fs::remove_file(&plan.candidate).with_context(|| {
                        format!(
                            "remove invalid wide target candidate {}",
                            plan.candidate.display()
                        )
                    })?;
                }
                return Err(error).context("stage=replacement_contract_validation");
            }
            stages.push(serde_json::json!({
                "name": "replacement_contract_validation",
                "status": "succeeded",
                "elapsed_ms": validation_started.elapsed().as_millis()
            }));
        }
        let mut output = serde_json::to_value(&execution)?;
        if staging.is_some() {
            output["replacement_validation"] = serde_json::Value::String("passed".to_owned());
        }
        if let Some(expected_target_hash) = expected_target_hash {
            let commit =
                replace_target_from_candidate(&plan.candidate, &plan.target, &expected_target_hash)
                    .context("stage=target_commit")?;
            stages.push(serde_json::json!({
                "name": "target_commit",
                "status": "succeeded",
                "elapsed_ms": commit.elapsed_ms
            }));
            output["target_commit"] = serde_json::to_value(commit)?;
        } else {
            stages.push(serde_json::json!({
                "name": "target_commit",
                "status": "not_requested"
            }));
        }
        output["execution_summary"] = serde_json::json!({
            "function": "append_target_execute",
            "status": "succeeded",
            "task_success_count": plan.task_ok,
            "task_skipped_count": plan.task_skipped,
            "appended_rows": plan.appended_rows,
            "duplicate_rows_filtered": plan.duplicate_rows_filtered,
            "elapsed_ms": run_started.elapsed().as_millis(),
            "stages": stages
        });
        println!("{}", serde_json::to_string_pretty(&output)?);
    } else {
        let write_started = Instant::now();
        let candidate =
            write_new_append_candidate(&plan).context("stage=new_workbook_write_and_reopen")?;
        let mut stages = vec![serde_json::json!({
            "name": "new_workbook_write_and_reopen",
            "status": "succeeded",
            "engine": "rust_xlsxwriter",
            "elapsed_ms": write_started.elapsed().as_millis()
        })];
        let mut output = serde_json::json!({
            "candidate": candidate,
            "candidate_engine": "rust_xlsxwriter",
            "status": "created_and_reopened"
        });
        if commit_target {
            let commit = create_target_from_candidate(&candidate, &plan.target)
                .context("stage=target_commit")?;
            stages.push(serde_json::json!({
                "name": "target_commit",
                "status": "succeeded",
                "elapsed_ms": commit.elapsed_ms
            }));
            output["target_commit"] = serde_json::to_value(commit)?;
            output["status"] = serde_json::Value::String("created_candidate_and_target".into());
        } else {
            stages.push(serde_json::json!({
                "name": "target_commit",
                "status": "not_requested"
            }));
        }
        output["execution_summary"] = serde_json::json!({
            "function": "append_target_execute",
            "status": "succeeded",
            "task_success_count": plan.task_ok,
            "task_skipped_count": plan.task_skipped,
            "appended_rows": plan.appended_rows,
            "duplicate_rows_filtered": plan.duplicate_rows_filtered,
            "elapsed_ms": run_started.elapsed().as_millis(),
            "stages": stages
        });
        println!("{}", serde_json::to_string_pretty(&output)?);
    }
    Ok(())
}

struct TemporaryFile(PathBuf);

impl TemporaryFile {
    fn new(path: PathBuf) -> Self {
        Self(path)
    }

    fn path(&self) -> &Path {
        &self.0
    }
}

impl Drop for TemporaryFile {
    fn drop(&mut self) {
        let _ = fs::remove_file(&self.0);
    }
}

fn run_apply_edit_plan(arguments: Vec<String>) -> Result<()> {
    let options = parse_cli_arguments(arguments)?;
    let plan = required_path(&options, "plan")?;
    let capabilities = required_path(&options, "capabilities")?;
    let platform = options
        .values
        .get("platform")
        .context("--platform is required")?;
    let report = apply_plan_file(plan, capabilities, platform)?;
    println!("{}", serde_json::to_string_pretty(&report)?);
    Ok(())
}

fn parse_duplicate_action(value: &str) -> Result<DuplicateAction> {
    match value {
        "check" => Ok(DuplicateAction::Check),
        "mark" => Ok(DuplicateAction::Mark),
        "delete" => Ok(DuplicateAction::Delete),
        _ => bail!("--action must be check, mark, or delete"),
    }
}

fn run_print_comment_plan(arguments: Vec<String>) -> Result<()> {
    let options = parse_cli_arguments(arguments)?;
    let input = required_path(&options, "input")?;
    let output = required_path(&options, "output")?;
    let drop_comments = options
        .values
        .get("drop-comments")
        .is_some_and(|value| parse_cli_bool(value));
    let hide_zero = options
        .values
        .get("hide-zero")
        .is_some_and(|value| parse_cli_bool(value));
    let report = build_comment_print_plan(&input, &output, drop_comments, hide_zero)?;
    let capabilities = options.values.get("capabilities");
    let platform = options.values.get("platform");
    anyhow::ensure!(
        capabilities.is_some() == platform.is_some(),
        "--capabilities and --platform must be provided together"
    );
    let mut result = serde_json::json!({
        "input": report.input,
        "output": report.output,
        "matching_sheets": report.matching_sheets,
        "marked_ranges": report.marked_ranges,
        "plan": report.plan,
        "execution": if capabilities.is_some() { "not_needed" } else { "plan_only" },
    });
    let planned = result.get("plan").filter(|plan| !plan.is_null()).cloned();
    if let Some(plan_value) = planned {
        if let Some(plan_path) = options.values.get("plan-output") {
            if let Some(parent) = Path::new(plan_path)
                .parent()
                .filter(|path| !path.as_os_str().is_empty())
            {
                fs::create_dir_all(parent)?;
            }
            fs::write(plan_path, serde_json::to_vec_pretty(&plan_value)?)?;
            result["plan_file"] = serde_json::json!(plan_path);
        }
        if let (Some(capabilities), Some(platform)) = (capabilities, platform) {
            let plan: excel_editor::ExcelEditPlan = serde_json::from_value(plan_value)?;
            let execution = apply_plan_with_capability_file(&plan, capabilities, platform)?;
            result["execution"] = serde_json::json!(execution);
        }
    }
    println!("{}", serde_json::to_string_pretty(&result)?);
    Ok(())
}

fn parse_cli_bool(value: &str) -> bool {
    matches!(
        value.trim().to_ascii_lowercase().as_str(),
        "1" | "true" | "yes" | "y" | "是" | "on"
    )
}

fn run_print_config_precheck(arguments: Vec<String>) -> Result<()> {
    let options = parse_cli_arguments(arguments)?;
    let config = required_path(&options, "config")?;
    let tasks = pytools_core::config::load_print_tasks(config)?;
    let report = precheck_print_tasks(&tasks);
    println!("{}", serde_json::to_string_pretty(&report)?);
    if report.failed > 0 {
        bail!("print precheck found {} failure(s)", report.failed);
    }
    Ok(())
}

fn run_print_config_plan(arguments: Vec<String>) -> Result<()> {
    let options = parse_cli_arguments(arguments)?;
    let config = required_path(&options, "config")?;
    let output_dir = match options.values.get("output-dir") {
        Some(path) => PathBuf::from(path),
        None => load_global(&config)?.output_dir,
    };
    let tasks = pytools_core::config::load_print_tasks(config)?;
    let reports = build_config_print_plans(&tasks)?;
    if !reports.is_empty() {
        fs::create_dir_all(&output_dir)?;
    }
    let capabilities = options.values.get("capabilities");
    let platform = options.values.get("platform");
    anyhow::ensure!(
        capabilities.is_some() == platform.is_some(),
        "--capabilities and --platform must be provided together"
    );
    let mut output = Vec::new();
    for (index, report) in reports.iter().enumerate() {
        let plan = report
            .plan
            .as_ref()
            .context("grouped print report has no edit plan")?;
        let path = output_dir.join(format!("print-target-{index}.plan.json"));
        fs::write(&path, serde_json::to_vec_pretty(plan)?)?;
        let execution = if let (Some(capabilities), Some(platform)) = (capabilities, platform) {
            let output_parent = plan
                .output
                .parent()
                .filter(|parent| !parent.as_os_str().is_empty())
                .unwrap_or_else(|| std::path::Path::new("."));
            fs::create_dir_all(output_parent).with_context(|| {
                format!("create print target directory {}", output_parent.display())
            })?;
            serde_json::json!(apply_plan_with_capability_file(
                plan,
                capabilities,
                platform
            )?)
        } else {
            serde_json::json!("plan_only")
        };
        output.push(serde_json::json!({
            "report": report,
            "plan_file": path,
            "execution": execution,
        }));
    }
    println!("{}", serde_json::to_string_pretty(&output)?);
    Ok(())
}

fn run_village_split_plan(arguments: Vec<String>) -> Result<()> {
    let options = parse_cli_arguments(arguments)?;
    let current = required_path(&options, "current")?;
    let previous = required_path(&options, "previous")?;
    let input = required_path(&options, "input")?;
    let output = required_path(&options, "output")?;
    let report = build_village_split_plan(&current, &previous, &input, &output)?;
    write_optional_plan(&options, &report.plan)?;
    let execution = execute_optional_plan(&options, report.plan.as_ref())?;
    let mut result = serde_json::to_value(report)?;
    result["execution"] = execution;
    println!("{}", serde_json::to_string_pretty(&result)?);
    Ok(())
}

fn run_village_split_batch(arguments: Vec<String>) -> Result<()> {
    let (commit_target, arguments) = extract_boolean_flag(arguments, "commit-target")?;
    let options = parse_cli_arguments(arguments)?;
    let current = required_path(&options, "current")?;
    let previous = required_path(&options, "previous")?;
    anyhow::ensure!(
        current
            .extension()
            .is_some_and(|extension| extension.eq_ignore_ascii_case("xlsx")),
        "PT-7-1 current-period input currently supports .xlsx only"
    );
    anyhow::ensure!(
        previous
            .extension()
            .is_some_and(|extension| extension.eq_ignore_ascii_case("xlsx")),
        "PT-7-1 prior-period input currently supports .xlsx only"
    );
    anyhow::ensure!(
        !options.positional.is_empty(),
        "provide at least one target workbook"
    );
    let capabilities = options.values.get("capabilities");
    let platform = options.values.get("platform");
    anyhow::ensure!(
        capabilities.is_some() == platform.is_some(),
        "--capabilities and --platform must be provided together"
    );
    anyhow::ensure!(
        !commit_target || capabilities.is_some(),
        "--commit-target requires --capabilities and --platform"
    );
    let log_view = options
        .values
        .get("log-view")
        .map(String::as_str)
        .unwrap_or("user");
    anyhow::ensure!(
        matches!(log_view, "user" | "developer" | "off"),
        "--log-view must be user, developer, or off"
    );
    let config = options
        .values
        .get("config")
        .map(|path| load_global(path))
        .transpose()?;
    let output_dir = match options.values.get("output-dir") {
        Some(path) => PathBuf::from(path),
        None => config
            .as_ref()
            .map(|config| config.output_dir.clone())
            .unwrap_or_else(|| PathBuf::from("output")),
    };
    let log_dir = match options.values.get("log-dir") {
        Some(path) => PathBuf::from(path),
        None => config
            .as_ref()
            .map(|config| config.log_dir.clone())
            .unwrap_or_else(|| PathBuf::from("logs")),
    };
    let plan_dir = unique_menu_action_dir(&output_dir, "village-split");
    fs::create_dir_all(&output_dir)
        .with_context(|| format!("create output directory {}", output_dir.display()))?;
    fs::create_dir_all(&plan_dir)
        .with_context(|| format!("create plan directory {}", plan_dir.display()))?;

    let started = Instant::now();
    let started_at = Local::now().to_rfc3339();
    let nonce = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map(|duration| duration.as_nanos())
        .unwrap_or_default();
    let run_id = format!("pytools-pt-7-1-{}-{nonce}", std::process::id());
    let mut planned_files = 0usize;
    let mut committed_files = 0usize;
    let mut completed_files = 0usize;
    let mut failed_files = 0usize;
    let mut matched_sheets = 0usize;
    let mut changed_cells = 0usize;
    let total = options.positional.len();
    let mut events = Vec::<SummaryRunEvent>::new();
    let mut progress = SummaryProgress {
        total_inputs: total,
        total_files: total,
        ..SummaryProgress::default()
    };
    let mut items = Vec::new();
    for (index, input) in options.positional.iter().enumerate() {
        let display_name = input
            .file_name()
            .map(|name| name.to_string_lossy().into_owned())
            .unwrap_or_else(|| input.display().to_string());
        progress.current_file = Some(index + 1);
        progress.processed_files = index;
        progress.completed_files = completed_files;
        progress.failed_files = failed_files;
        if log_view == "user" {
            eprintln!(
                "[处理中] {}/{} 拆分村镇银行数据：{}",
                index + 1,
                total,
                display_name
            );
        }
        emit_village_split_event(
            &mut events,
            log_view,
            started,
            "planning",
            "started",
            format!("开始处理第 {}/{} 个工作簿", index + 1, total),
            &progress,
            None,
            &["pytools-core"],
            None,
            None,
        );
        let file_started = Instant::now();
        let mut current_stage = "planning";
        let candidate = unique_workbook_candidate(input, &output_dir, "village_split");
        let plan_path = plan_dir.join(format!("target-{:04}.plan.json", index + 1));
        let item_result = (|| -> Result<serde_json::Value> {
            anyhow::ensure!(
                input
                    .extension()
                    .is_some_and(|extension| extension.eq_ignore_ascii_case("xlsx")),
                "PT-7-1 target currently supports .xlsx only"
            );
            let expected_target_hash = if commit_target {
                Some(sha256_file(input)?)
            } else {
                None
            };
            let planning_started = Instant::now();
            let report = build_village_split_plan(&current, &previous, input, &candidate)?;
            if let Some(expected_hash) = expected_target_hash.as_deref() {
                verify_planning_snapshot(input, Some(expected_hash))?;
            }
            emit_village_split_event(
                &mut events,
                log_view,
                started,
                "planning",
                if report.plan.is_some() {
                    "completed"
                } else {
                    "no_changes"
                },
                if report.plan.is_some() {
                    "已完成业务计算并生成工作簿修改计划"
                } else {
                    "没有需要修改的工作表"
                },
                &progress,
                Some(planning_started),
                &["pytools-core", "excel-reader", "calamine"],
                Some("calamine"),
                None,
            );
            let execution = if let Some(plan) = report.plan.as_ref() {
                current_stage = "plan_persistence";
                let persistence_started = Instant::now();
                fs::write(&plan_path, serde_json::to_vec_pretty(plan)?)
                    .with_context(|| format!("write edit plan {}", plan_path.display()))?;
                emit_village_split_event(
                    &mut events,
                    log_view,
                    started,
                    "plan_persistence",
                    "completed",
                    "已保存工作簿修改计划",
                    &progress,
                    Some(persistence_started),
                    &["pytools-core"],
                    None,
                    None,
                );
                current_stage = "engine_execution";
                let execution_started = Instant::now();
                let execution = execute_optional_plan(&options, Some(plan))?;
                let engine = execution.get("engine").and_then(serde_json::Value::as_str);
                let plan_only = execution.as_str() == Some("plan_only");
                emit_village_split_event(
                    &mut events,
                    log_view,
                    started,
                    "engine_execution",
                    if plan_only { "skipped" } else { "completed" },
                    if plan_only {
                        "已生成计划；当前没有已验收的执行 capability，原工作簿未修改"
                    } else {
                        "已完成工作簿候选修改与引擎验证"
                    },
                    &progress,
                    Some(execution_started),
                    if plan_only {
                        &["pytools-core"]
                    } else {
                        &["pytools-core", "excel-editor"]
                    },
                    engine,
                    None,
                );
                execution
            } else {
                serde_json::Value::String("no_changes".to_owned())
            };
            let commit = if commit_target && report.plan.is_some() {
                current_stage = "safe_commit";
                let commit_started = Instant::now();
                let commit = replace_target_from_candidate(
                    &candidate,
                    input,
                    expected_target_hash
                        .as_deref()
                        .context("target snapshot missing before commit")?,
                )?;
                emit_village_split_event(
                    &mut events,
                    log_view,
                    started,
                    "safe_commit",
                    "completed",
                    "已验证候选并安全提交到目标工作簿",
                    &progress,
                    Some(commit_started),
                    &["pytools-core", "file-commit"],
                    execution.get("engine").and_then(serde_json::Value::as_str),
                    None,
                );
                Some(commit)
            } else {
                None
            };
            Ok(serde_json::json!({
                "input": input,
                "candidate": candidate,
                "plan_file": if report.plan.is_some() { Some(plan_path) } else { None },
                "status": if commit.is_some() {
                    "committed"
                } else if report.plan.is_none() {
                    "no_changes"
                } else if execution.as_str() == Some("plan_only") {
                    "planned"
                } else {
                    "candidate_created"
                },
                "matched_sheets": report.matched_sheets,
                "changed_cells": report.changed_cells,
                "warnings": report.warnings,
                "execution": execution,
                "target_commit": commit,
                "elapsed_ms": file_started.elapsed().as_millis(),
            }))
        })();
        match item_result {
            Ok(item) => {
                if item["plan_file"].is_string() {
                    planned_files += 1;
                }
                if item["status"] == "committed" {
                    committed_files += 1;
                }
                completed_files += 1;
                matched_sheets += item["matched_sheets"].as_u64().unwrap_or_default() as usize;
                changed_cells += item["changed_cells"].as_u64().unwrap_or_default() as usize;
                let (status, message) = match item["status"].as_str().unwrap_or_default() {
                    "committed" => ("completed", "当前工作簿已安全提交"),
                    "planned" => ("planned", "当前工作簿修改计划已生成，原文件未修改"),
                    "candidate_created" => ("completed", "当前工作簿候选已生成"),
                    _ => ("no_changes", "当前工作簿无需修改"),
                };
                items.push(item);
                progress.processed_files = index + 1;
                progress.completed_files = completed_files;
                emit_village_split_event(
                    &mut events,
                    log_view,
                    started,
                    "file_summary",
                    status,
                    message,
                    &progress,
                    Some(file_started),
                    &["pytools-core"],
                    None,
                    None,
                );
            }
            Err(error) => {
                failed_files += 1;
                progress.processed_files = index + 1;
                progress.failed_files = failed_files;
                emit_village_split_event(
                    &mut events,
                    log_view,
                    started,
                    current_stage,
                    "failed",
                    "当前工作簿处理失败；请查看开发者运行记录中的失败阶段",
                    &progress,
                    Some(file_started),
                    &["pytools-core"],
                    None,
                    Some("village_split_stage_error"),
                );
                items.push(serde_json::json!({
                    "input": input,
                    "status": "failed",
                    "error_stage": current_stage,
                    "error": format!("{error:#}"),
                    "elapsed_ms": file_started.elapsed().as_millis(),
                }));
            }
        }
    }
    let status = if failed_files == 0 {
        "succeeded"
    } else if completed_files == 0 {
        "failed"
    } else {
        "partial"
    };
    let mut result = serde_json::json!({
        "feature_id": "PT-7-1",
        "files": total,
        "planned_files": planned_files,
        "completed_files": completed_files,
        "committed_files": committed_files,
        "failed_files": failed_files,
        "matched_sheets": matched_sheets,
        "changed_cells": changed_cells,
        "mode": if commit_target { "commit_target" } else if capabilities.is_some() { "candidate" } else { "plan_only" },
        "status": status,
        "plan_dir": plan_dir,
        "elapsed_ms": started.elapsed().as_millis(),
        "items": items,
    });
    let run_log_path = if log_view == "off" {
        None
    } else {
        let path = log_dir.join(format!("PT-7-1_{run_id}.run.json"));
        let safe_items = result["items"]
            .as_array()
            .into_iter()
            .flatten()
            .map(|item| {
                serde_json::json!({
                    "status": item["status"],
                    "error_stage": item.get("error_stage"),
                    "error_type": if item["status"] == "failed" {
                        Some("village_split_stage_error")
                    } else {
                        None
                    },
                    "matched_sheets": item.get("matched_sheets"),
                    "changed_cells": item.get("changed_cells"),
                    "elapsed_ms": item.get("elapsed_ms"),
                })
            })
            .collect::<Vec<_>>();
        let history = serde_json::json!({
            "schema_version": 1,
            "product": "pytools",
            "feature_id": "PT-7-1",
            "function": "village-split",
            "run_id": run_id,
            "started_at": started_at,
            "status": status,
            "total_inputs": total,
            "completed_files": completed_files,
            "failed_files": failed_files,
            "planned_files": planned_files,
            "committed_files": committed_files,
            "matched_sheets": matched_sheets,
            "changed_cells": changed_cells,
            "elapsed_ms": started.elapsed().as_millis(),
            "items": safe_items,
            "events": events,
        });
        match write_json_report_create_new(&path, &history) {
            Ok(()) => {
                result["run_id"] = serde_json::Value::String(run_id.clone());
                result["run_log_path"] = serde_json::Value::String(path.display().to_string());
                Some(path)
            }
            Err(error) => {
                eprintln!("pytools: warning: could not save PT-7-1 run history: {error:#}");
                None
            }
        }
    };
    if log_view == "user" {
        if let Some(path) = run_log_path {
            eprintln!("运行记录已保存：{}", path.display());
        }
    }
    println!("{}", serde_json::to_string_pretty(&result)?);
    Ok(())
}

fn emit_village_split_event(
    events: &mut Vec<SummaryRunEvent>,
    log_view: &str,
    run_started: Instant,
    phase: &str,
    status: &str,
    user_message: impl Into<String>,
    progress: &SummaryProgress,
    stage_started: Option<Instant>,
    components: &[&str],
    actual_engine: Option<&str>,
    error_type: Option<&str>,
) {
    if log_view == "off" {
        return;
    }
    let event = SummaryRunEvent {
        sequence: events.len() + 1,
        timestamp: Local::now().to_rfc3339(),
        phase: phase.to_owned(),
        status: status.to_owned(),
        user_message: user_message.into(),
        progress: progress.clone(),
        file_name: None,
        stage_elapsed_ms: stage_started.map(|started| started.elapsed().as_millis()),
        total_elapsed_ms: run_started.elapsed().as_millis(),
        components: components
            .iter()
            .map(|component| (*component).to_owned())
            .collect(),
        actual_engine: actual_engine.map(str::to_owned),
        error_type: error_type.map(str::to_owned),
    };
    if log_view == "developer" {
        if let Ok(encoded) = serde_json::to_string(&event) {
            eprintln!("{encoded}");
        } else {
            eprintln!("{{\"phase\":\"progress_event\",\"status\":\"serialization_failed\"}}");
        }
    } else {
        eprintln!("{}", event.user_message);
    }
    events.push(event);
}

fn run_remove_foreign_bank_plan(arguments: Vec<String>) -> Result<()> {
    let options = parse_cli_arguments(arguments)?;
    let config = required_path(&options, "config")?;
    let input = required_path(&options, "input")?;
    let output = required_path(&options, "output")?;
    let foreign = pytools_core::config::load_foreign_banks(config)?;
    let report = build_remove_foreign_plan(&input, &output, &foreign)?;
    write_optional_plan(&options, &report.plan)?;
    let execution = execute_optional_plan(&options, report.plan.as_ref())?;
    let mut result = serde_json::to_value(report)?;
    result["execution"] = execution;
    println!("{}", serde_json::to_string_pretty(&result)?);
    Ok(())
}

fn run_remove_foreign_bank_batch(arguments: Vec<String>) -> Result<()> {
    let options = parse_cli_arguments(arguments)?;
    let config = required_path(&options, "config")?;
    anyhow::ensure!(
        !options.positional.is_empty(),
        "provide at least one source workbook"
    );
    let capabilities = options.values.get("capabilities");
    let platform = options.values.get("platform");
    anyhow::ensure!(
        capabilities.is_some() == platform.is_some(),
        "--capabilities and --platform must be provided together"
    );
    let execute = capabilities.is_some();
    let view = options.values.get("view").map(String::as_str);
    if let Some(view) = view {
        anyhow::ensure!(
            matches!(view, "user" | "developer" | "off"),
            "--view must be user, developer, or off"
        );
    }
    anyhow::ensure!(
        view.is_some() || !options.values.contains_key("report-output"),
        "--report-output requires --view user or developer"
    );
    let plan_dir = options.values.get("plan-dir").map(PathBuf::from);
    anyhow::ensure!(
        !execute || plan_dir.is_none(),
        "--plan-dir is only supported in plan-only mode"
    );
    let report_output = if view == Some("off") {
        anyhow::ensure!(
            !options.values.contains_key("report-output"),
            "--report-output cannot be used with --view off"
        );
        None
    } else if view.is_some() {
        let path = options
            .values
            .get("report-output")
            .map(PathBuf::from)
            .unwrap_or_else(|| {
                let nonce = SystemTime::now()
                    .duration_since(UNIX_EPOCH)
                    .map(|duration| duration.as_nanos())
                    .unwrap_or_default();
                let log_dir = options
                    .values
                    .get("log-dir")
                    .map(PathBuf::from)
                    .unwrap_or_else(|| PathBuf::from("logs"));
                log_dir.join(format!(
                    "PT-7-3_{}_{}_{}.json",
                    Local::now().format("%Y%m%dT%H%M%S"),
                    std::process::id(),
                    nonce
                ))
            });
        anyhow::ensure!(
            path.extension()
                .and_then(|extension| extension.to_str())
                .is_some_and(|extension| extension.eq_ignore_ascii_case("json")),
            "--report-output must use a .json extension"
        );
        anyhow::ensure!(
            !path.exists(),
            "run history file already exists; refusing to overwrite {}",
            path.display()
        );
        Some(path)
    } else {
        None
    };

    let run_started = Instant::now();
    let started_at = Local::now().to_rfc3339();
    let nonce = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map(|duration| duration.as_nanos())
        .unwrap_or_default();
    let run_id = format!("pytools-pt-7-3-{}-{nonce}", std::process::id());
    let configuration_started = Instant::now();
    let foreign = pytools_core::config::load_foreign_banks(&config)?;
    let configuration_elapsed_ms = configuration_started.elapsed().as_millis();
    let mut result = serde_json::json!({
        "files": 0,
        "ok_files": 0,
        "branch_sheets": 0,
        "deleted_rows": 0,
        "failed_files": 0,
        "planned_files": 0,
        "planned_branch_sheets": 0,
        "planned_deleted_rows": 0,
        "foreign_empty": foreign.is_empty(),
        "mode": if execute { "execute" } else { "plan_only" },
        "items": [],
    });

    let mut events = Vec::<SummaryRunEvent>::new();
    let mut progress = SummaryProgress {
        total_inputs: options.positional.len(),
        total_files: options.positional.len(),
        ..SummaryProgress::default()
    };
    if view.is_some() {
        emit_foreign_bank_event(
            &mut events,
            view,
            run_started,
            "configuration_load",
            "completed",
            "外资行配置读取完成。",
            &progress,
            Some(configuration_elapsed_ms),
            &["excel-reader", "calamine"],
            Some("calamine"),
            None,
        );
    }

    if foreign.is_empty() {
        if view.is_some() {
            emit_foreign_bank_event(
                &mut events,
                view,
                run_started,
                "configuration_precheck",
                "skipped",
                "外资行配置为空，未处理工作簿。",
                &progress,
                None,
                &["pytools-core::config"],
                None,
                Some("empty_foreign_bank_mapping"),
            );
        }
    } else {
        result["files"] = serde_json::json!(options.positional.len());
        for (file_index, input) in options.positional.into_iter().enumerate() {
            progress.current_file = Some(file_index + 1);
            progress.processed_files = file_index + 1;
            if view.is_some() {
                emit_foreign_bank_event(
                    &mut events,
                    view,
                    run_started,
                    "file_start",
                    "started",
                    format!(
                        "开始检查第 {}/{} 个工作簿。",
                        file_index + 1,
                        progress.total_files
                    ),
                    &progress,
                    None,
                    &[],
                    None,
                    None,
                );
            }
            let started = Instant::now();
            let output = match foreign_bank_copy_path(&input) {
                Ok(path) => path,
                Err(error) => {
                    result["failed_files"] =
                        serde_json::json!(result["failed_files"].as_u64().unwrap_or_default() + 1);
                    result["items"]
                        .as_array_mut()
                        .unwrap()
                        .push(serde_json::json!({
                            "input": input,
                            "status": "failed",
                            "error_stage": "input_validation",
                            "error": format!("{error:#}"),
                            "elapsed_ms": started.elapsed().as_millis(),
                        }));
                    progress.failed_files += 1;
                    if view.is_some() {
                        emit_foreign_bank_event(
                            &mut events,
                            view,
                            run_started,
                            "input_validation",
                            "failed",
                            "工作簿路径不符合处理要求。",
                            &progress,
                            Some(started.elapsed().as_millis()),
                            &[],
                            None,
                            Some("input_validation"),
                        );
                    }
                    continue;
                }
            };
            let extension_is_xlsx = input
                .extension()
                .and_then(|value| value.to_str())
                .is_some_and(|value| value.eq_ignore_ascii_case("xlsx"));
            if !extension_is_xlsx {
                result["failed_files"] =
                    serde_json::json!(result["failed_files"].as_u64().unwrap_or_default() + 1);
                result["items"]
                    .as_array_mut()
                    .unwrap()
                    .push(serde_json::json!({
                        "input": input,
                        "output": output,
                        "status": "failed",
                        "error_stage": "input_validation",
                        "error": "Rust row-removal candidate currently accepts .xlsx only",
                        "elapsed_ms": started.elapsed().as_millis(),
                    }));
                progress.failed_files += 1;
                if view.is_some() {
                    emit_foreign_bank_event(
                        &mut events,
                        view,
                        run_started,
                        "input_validation",
                        "failed",
                        "输入格式尚未验收；本项只接受 .xlsx。",
                        &progress,
                        Some(started.elapsed().as_millis()),
                        &[],
                        None,
                        Some("unsupported_format"),
                    );
                }
                continue;
            }
            let plan_started = Instant::now();
            let report = match build_remove_foreign_plan(&input, &output, &foreign) {
                Ok(report) => report,
                Err(error) => {
                    result["failed_files"] =
                        serde_json::json!(result["failed_files"].as_u64().unwrap_or_default() + 1);
                    result["items"]
                        .as_array_mut()
                        .unwrap()
                        .push(serde_json::json!({
                            "input": input,
                            "output": output,
                            "status": "failed",
                            "error_stage": "planning",
                            "error": format!("{error:#}"),
                            "elapsed_ms": started.elapsed().as_millis(),
                        }));
                    progress.failed_files += 1;
                    if view.is_some() {
                        emit_foreign_bank_event(
                            &mut events,
                            view,
                            run_started,
                            "plan_build",
                            "failed",
                            "工作簿检查或修改计划生成失败。",
                            &progress,
                            Some(plan_started.elapsed().as_millis()),
                            &["pytools-core::batch_ops", "excel-reader"],
                            None,
                            Some("planning"),
                        );
                    }
                    continue;
                }
            };
            let plan_elapsed_ms = plan_started.elapsed().as_millis();
            let planned_deleted_rows = report
                .plan
                .as_ref()
                .map(|plan| {
                    plan.operations
                        .iter()
                        .map(|operation| match operation {
                            excel_editor::EditOperation::DeleteRows { count, .. } => {
                                *count as usize
                            }
                            _ => 0,
                        })
                        .sum::<usize>()
                })
                .unwrap_or_default();
            result["planned_files"] =
                serde_json::json!(result["planned_files"].as_u64().unwrap_or_default() + 1);
            result["planned_branch_sheets"] = serde_json::json!(
                result["planned_branch_sheets"].as_u64().unwrap_or_default()
                    + report.matched_sheets as u64
            );
            result["planned_deleted_rows"] = serde_json::json!(
                result["planned_deleted_rows"].as_u64().unwrap_or_default()
                    + planned_deleted_rows as u64
            );

            let mut item = serde_json::json!({
                "input": input,
                "output": output,
                "status": if execute { "pending" } else if report.plan.is_some() { "plan_only" } else { "plan_only_no_changes" },
                "planned_branch_sheets": report.matched_sheets,
                "planned_deleted_rows": planned_deleted_rows,
                "branch_sheets": 0,
                "deleted_rows": 0,
                "plan": report.plan.clone(),
            });

            if !execute && !(plan_dir.is_some() && report.plan.is_some()) {
                progress.completed_files += 1;
            }
            if view.is_some() {
                emit_foreign_bank_event(
                    &mut events,
                    view,
                    run_started,
                    "plan_build",
                    "completed",
                    format!(
                        "修改计划已生成：{} 个目标工作表，{} 行待删除。",
                        report.matched_sheets, planned_deleted_rows
                    ),
                    &progress,
                    Some(plan_elapsed_ms),
                    &["pytools-core::batch_ops", "excel-reader", "ExcelEditPlan"],
                    None,
                    None,
                );
            }

            if let (Some(plan_dir), Some(plan)) = (plan_dir.as_deref(), report.plan.as_ref()) {
                let plan_path = foreign_bank_plan_path(plan_dir, file_index, &input);
                let persistence_started = Instant::now();
                match write_json_report_create_new(&plan_path, plan) {
                    Ok(()) => {
                        item["plan_path"] = serde_json::json!(plan_path);
                        progress.completed_files += 1;
                        if view.is_some() {
                            emit_foreign_bank_event(
                                &mut events,
                                view,
                                run_started,
                                "plan_persistence",
                                "completed",
                                "修改计划已保存。",
                                &progress,
                                Some(persistence_started.elapsed().as_millis()),
                                &["excel-editor::ExcelEditPlan"],
                                None,
                                None,
                            );
                        }
                    }
                    Err(error) => {
                        result["failed_files"] = serde_json::json!(
                            result["failed_files"].as_u64().unwrap_or_default() + 1
                        );
                        item["status"] = serde_json::json!("failed");
                        item["error_stage"] = serde_json::json!("plan_persistence");
                        item["error"] = serde_json::json!(format!("{error:#}"));
                        item["elapsed_ms"] = serde_json::json!(started.elapsed().as_millis());
                        progress.failed_files += 1;
                        if view.is_some() {
                            emit_foreign_bank_event(
                                &mut events,
                                view,
                                run_started,
                                "plan_persistence",
                                "failed",
                                "修改计划保存失败。",
                                &progress,
                                Some(persistence_started.elapsed().as_millis()),
                                &["excel-editor::ExcelEditPlan"],
                                None,
                                Some("plan_persistence"),
                            );
                        }
                        result["items"].as_array_mut().unwrap().push(item);
                        continue;
                    }
                }
            }

            if !execute {
                item["execution"] = serde_json::json!("plan_only");
            } else {
                let execution_started = Instant::now();
                let execution: Result<serde_json::Value> = if let Some(plan) =
                    item.get("plan").filter(|value| !value.is_null()).cloned()
                {
                    let plan: excel_editor::ExcelEditPlan = serde_json::from_value(plan)?;
                    apply_plan_with_capability_file(
                        &plan,
                        capabilities.expect("validated capability path"),
                        platform.expect("validated platform"),
                    )
                    .and_then(|report| {
                        serde_json::to_value(report).context("serialize edit result")
                    })
                } else {
                    copy_workbook_to_new_target(&input, &output).and_then(|report| {
                        serde_json::to_value(report).context("serialize copy result")
                    })
                };
                match execution {
                    Ok(execution) => {
                        result["ok_files"] =
                            serde_json::json!(result["ok_files"].as_u64().unwrap_or_default() + 1);
                        result["branch_sheets"] = serde_json::json!(
                            result["branch_sheets"].as_u64().unwrap_or_default()
                                + report.matched_sheets as u64
                        );
                        result["deleted_rows"] = serde_json::json!(
                            result["deleted_rows"].as_u64().unwrap_or_default()
                                + planned_deleted_rows as u64
                        );
                        item["status"] = serde_json::json!("written");
                        item["branch_sheets"] = serde_json::json!(report.matched_sheets);
                        item["deleted_rows"] = serde_json::json!(planned_deleted_rows);
                        let actual_engine = foreign_bank_actual_engine(&execution);
                        item["execution"] = execution;
                        progress.completed_files += 1;
                        if view.is_some() {
                            emit_foreign_bank_event(
                                &mut events,
                                view,
                                run_started,
                                "engine_execution",
                                "completed",
                                "外发副本已生成，原工作簿保持不变。",
                                &progress,
                                Some(execution_started.elapsed().as_millis()),
                                &["ExcelModifier"],
                                actual_engine.as_deref(),
                                None,
                            );
                        }
                    }
                    Err(error) => {
                        result["failed_files"] = serde_json::json!(
                            result["failed_files"].as_u64().unwrap_or_default() + 1
                        );
                        item["status"] = serde_json::json!("failed");
                        item["error_stage"] = serde_json::json!("execution");
                        item["error"] = serde_json::json!(format!("{error:#}"));
                        progress.failed_files += 1;
                        if view.is_some() {
                            emit_foreign_bank_event(
                                &mut events,
                                view,
                                run_started,
                                "engine_execution",
                                "failed",
                                "ExcelModifier 执行失败；未发布外发副本。",
                                &progress,
                                Some(execution_started.elapsed().as_millis()),
                                &["ExcelModifier"],
                                None,
                                Some("execution"),
                            );
                        }
                    }
                }
            }
            item["elapsed_ms"] = serde_json::json!(started.elapsed().as_millis());
            result["items"].as_array_mut().unwrap().push(item);
        }
    }

    let total_elapsed_ms = run_started.elapsed().as_millis();
    if let Some(view) = view {
        let user_view = foreign_bank_user_view(&result, total_elapsed_ms, Some(&events));
        if let Some(path) = report_output.as_deref() {
            let history = serde_json::json!({
                "schema_version": 1,
                "product": "pytools",
                "feature_id": "PT-7-3",
                "run_id": &run_id,
                "started_at": &started_at,
                "status": user_view["status"],
                "mode": result["mode"],
                "summary": user_view["summary"],
                "events": &events,
            });
            write_json_report_create_new(path, &history)?;
        }
        let output = match view {
            "developer" => {
                result["elapsed_ms"] = serde_json::json!(total_elapsed_ms);
                result["run_id"] = serde_json::json!(run_id);
                result["started_at"] = serde_json::json!(started_at);
                result["events"] = serde_json::to_value(&events)?;
                if let Some(path) = report_output {
                    result["history_report"] = serde_json::json!(path);
                }
                result
            }
            "user" => {
                let mut output = user_view;
                output["run_id"] = serde_json::json!(run_id);
                output["started_at"] = serde_json::json!(started_at);
                if let Some(path) = report_output {
                    output["history_report"] = serde_json::json!(path);
                }
                output
            }
            "off" => foreign_bank_user_view(&result, total_elapsed_ms, None),
            _ => unreachable!("view was validated"),
        };
        println!("{}", serde_json::to_string_pretty(&output)?);
    } else {
        println!("{}", serde_json::to_string_pretty(&result)?);
    }
    Ok(())
}

fn emit_foreign_bank_event(
    events: &mut Vec<SummaryRunEvent>,
    view: Option<&str>,
    run_started: Instant,
    phase: &str,
    status: &str,
    user_message: impl Into<String>,
    progress: &SummaryProgress,
    stage_elapsed_ms: Option<u128>,
    components: &[&str],
    actual_engine: Option<&str>,
    error_type: Option<&str>,
) {
    let Some(view) = view.filter(|view| *view != "off") else {
        return;
    };
    let event = SummaryRunEvent {
        sequence: events.len() + 1,
        timestamp: Local::now().to_rfc3339(),
        phase: phase.to_owned(),
        status: status.to_owned(),
        user_message: user_message.into(),
        progress: progress.clone(),
        file_name: None,
        stage_elapsed_ms,
        total_elapsed_ms: run_started.elapsed().as_millis(),
        components: components.iter().map(|value| (*value).to_owned()).collect(),
        actual_engine: actual_engine.map(str::to_owned),
        error_type: error_type.map(str::to_owned),
    };
    if view == "developer" {
        if let Ok(encoded) = serde_json::to_string(&event) {
            eprintln!("{encoded}");
        }
    } else {
        eprintln!(
            "{}（进度 {}/{}, 已完成 {}，失败 {}）",
            event.user_message,
            event.progress.processed_files,
            event.progress.total_files,
            event.progress.completed_files,
            event.progress.failed_files
        );
    }
    events.push(event);
}

fn foreign_bank_plan_path(plan_dir: &Path, file_index: usize, input: &Path) -> PathBuf {
    let stem = input
        .file_stem()
        .map(|stem| stem.to_string_lossy())
        .unwrap_or_default();
    let safe_stem = stem
        .chars()
        .map(|character| {
            if character.is_ascii_alphanumeric() || matches!(character, '-' | '_') {
                character
            } else {
                '_'
            }
        })
        .collect::<String>();
    plan_dir.join(format!("{:03}_{safe_stem}.json", file_index + 1))
}

fn foreign_bank_actual_engine(execution: &serde_json::Value) -> Option<String> {
    execution
        .get("engine")
        .and_then(serde_json::Value::as_str)
        .map(str::to_owned)
        .or_else(|| {
            (execution
                .get("operation")
                .and_then(serde_json::Value::as_str)
                == Some("copied"))
            .then(|| "filesystem_copy".to_owned())
        })
}

fn foreign_bank_user_view(
    result: &serde_json::Value,
    elapsed_ms: u128,
    events: Option<&[SummaryRunEvent]>,
) -> serde_json::Value {
    let is_plan_only = result["mode"] == "plan_only";
    let foreign_empty = result["foreign_empty"].as_bool().unwrap_or(false);
    let failed_files = result["failed_files"].as_u64().unwrap_or_default();
    let status = if foreign_empty {
        "not_run"
    } else if failed_files > 0 {
        "completed_with_errors"
    } else {
        "completed"
    };
    let completed_files = if is_plan_only {
        serde_json::json!(result["items"]
            .as_array()
            .map(|items| {
                items
                    .iter()
                    .filter(|item| {
                        matches!(
                            item["status"].as_str(),
                            Some("plan_only" | "plan_only_no_changes")
                        )
                    })
                    .count()
            })
            .unwrap_or_default())
    } else {
        result["ok_files"].clone()
    };
    let mut output = serde_json::json!({
        "schema_version": 1,
        "product": "pytools",
        "feature_id": "PT-7-3",
        "status": status,
        "mode": if is_plan_only { "plan_only" } else { "execute_copy" },
        "summary": {
            "total_files": result["files"],
            "completed_files": completed_files,
            "failed_files": result["failed_files"],
            "planned_files": result["planned_files"],
            "planned_branch_sheets": result["planned_branch_sheets"],
            "planned_deleted_rows": result["planned_deleted_rows"],
            "branch_sheets": result["branch_sheets"],
            "deleted_rows": result["deleted_rows"],
            "foreign_empty": foreign_empty,
            "elapsed_ms": elapsed_ms
        }
    });
    if let Some(events) = events {
        output["events"] = serde_json::Value::Array(
            events
                .iter()
                .map(|event| {
                    serde_json::json!({
                        "sequence": event.sequence,
                        "timestamp": event.timestamp,
                        "phase": foreign_bank_user_phase(&event.phase),
                        "status": foreign_bank_user_status(&event.status),
                        "message": event.user_message,
                        "progress": event.progress,
                        "elapsed_ms": event.stage_elapsed_ms,
                        "total_elapsed_ms": event.total_elapsed_ms
                    })
                })
                .collect(),
        );
    }
    output
}

fn foreign_bank_user_phase(phase: &str) -> &str {
    match phase {
        "configuration_load" => "配置读取",
        "configuration_precheck" => "配置检查",
        "file_start" | "input_validation" => "输入检查",
        "plan_build" => "修改计划",
        "plan_persistence" => "计划保存",
        "engine_execution" => "副本生成",
        other => other,
    }
}

fn foreign_bank_user_status(status: &str) -> &str {
    match status {
        "started" => "进行中",
        "completed" => "完成",
        "failed" => "失败",
        "skipped" => "跳过",
        other => other,
    }
}

fn foreign_bank_copy_path(input: &Path) -> Result<PathBuf> {
    let stem = input
        .file_stem()
        .with_context(|| format!("source workbook has no file stem: {}", input.display()))?;
    let extension = input
        .extension()
        .with_context(|| format!("source workbook has no extension: {}", input.display()))?;
    let mut file_name = stem.to_os_string();
    file_name.push("(金融局).");
    file_name.push(extension);
    Ok(input.with_file_name(file_name))
}

fn run_rural_loan_plan(arguments: Vec<String>) -> Result<()> {
    let options = parse_cli_arguments(arguments)?;
    let current = required_path(&options, "current")?;
    let previous = required_path(&options, "previous")?;
    let output = required_path(&options, "output")?;
    let (report, summary) = build_rural_loan_plan(&current, &previous, &output)?;
    write_optional_plan(&options, &report.plan)?;
    let execution = execute_optional_plan(&options, report.plan.as_ref())?;
    println!(
        "{}",
        serde_json::to_string_pretty(
            &serde_json::json!({"edit":report,"summary":summary,"execution":execution})
        )?
    );
    Ok(())
}

fn run_quarterly_rural_loan_plan(arguments: Vec<String>) -> Result<()> {
    let options = parse_cli_arguments(arguments)?;
    let current = required_path(&options, "current")?;
    let previous = required_path(&options, "previous")?;
    let output = required_path(&options, "output")?;
    let (report, summary) = build_quarterly_rural_loan_plan(&current, &previous, &output)?;
    write_optional_plan(&options, &report.plan)?;
    let execution = execute_optional_plan(&options, report.plan.as_ref())?;
    println!(
        "{}",
        serde_json::to_string_pretty(
            &serde_json::json!({"edit":report,"summary":summary,"execution":execution})
        )?
    );
    Ok(())
}

fn run_institution_normalize_plan(arguments: Vec<String>) -> Result<()> {
    let options = parse_cli_arguments(arguments)?;
    let config = required_path(&options, "config")?;
    let input = required_path(&options, "input")?;
    let output = required_path(&options, "output")?;
    let profile = excel_editor::inspect_xlsx_profile(&input)?;
    let mapping = pytools_core::config::load_institution_mapping_ordered(config)?;
    let report = pytools_core::institution_normalize::build_institution_normalize_plan(
        &input,
        &output,
        profile,
        "pytools-institution-normalize",
        &mapping,
    )?;
    write_optional_plan(&options, &report.plan)?;
    let execution = execute_optional_plan(&options, report.plan.as_ref())?;
    println!(
        "{}",
        serde_json::to_string_pretty(&serde_json::json!({
            "branch_sheets": report.branch_sheets,
            "mapped_sheets": report.mapped_sheets,
            "mapped_rows": report.mapped_rows,
            "mapping_empty": report.mapping_empty,
            "plan_written": report.plan.is_some(),
            "execution": execution,
        }))?
    );
    Ok(())
}

fn run_institution_normalize_batch(arguments: Vec<String>) -> Result<()> {
    let (apply, arguments) = extract_boolean_flag(arguments, "apply")?;
    let options = parse_cli_arguments(arguments)?;
    let config = required_path(&options, "config")?;
    let view = options
        .values
        .get("view")
        .map(String::as_str)
        .unwrap_or("user");
    anyhow::ensure!(
        matches!(view, "user" | "developer" | "off"),
        "--view must be user, developer, or off"
    );
    let plan_dir = options.values.get("plan-dir").map(PathBuf::from);
    let report_output = if view == "off" {
        anyhow::ensure!(
            !options.values.contains_key("report-output"),
            "--report-output cannot be used with --view off"
        );
        None
    } else {
        let path = options
            .values
            .get("report-output")
            .map(PathBuf::from)
            .unwrap_or_else(|| {
                let nonce = std::time::SystemTime::now()
                    .duration_since(std::time::UNIX_EPOCH)
                    .map(|duration| duration.as_nanos())
                    .unwrap_or_default();
                PathBuf::from("logs").join(format!(
                    "PT-7-2_{}_{}_{}.json",
                    chrono::Local::now().format("%Y%m%dT%H%M%S"),
                    std::process::id(),
                    nonce
                ))
            });
        anyhow::ensure!(
            path.extension()
                .and_then(|extension| extension.to_str())
                .is_some_and(|extension| extension.eq_ignore_ascii_case("json")),
            "--report-output must use a .json extension"
        );
        anyhow::ensure!(
            !path.exists(),
            "run history file already exists; refusing to overwrite {}",
            path.display()
        );
        if let Some(parent) = path
            .parent()
            .filter(|parent| !parent.as_os_str().is_empty())
        {
            fs::create_dir_all(parent)
                .with_context(|| format!("create history directory {}", parent.display()))?;
        }
        Some(path)
    };
    let capability = options.values.get("capabilities").map(PathBuf::from);
    let platform = options.values.get("platform").cloned();
    anyhow::ensure!(
        apply == capability.is_some() && apply == platform.is_some(),
        "--apply requires both --capabilities and --platform; omit all three for plan-only mode"
    );
    let run_options = pytools_core::institution_normalize::InstitutionNormalizeRunOptions {
        apply,
        capabilities: capability,
        platform,
        plan_dir,
    };
    let mut emitted_events = Vec::new();
    let report = pytools_core::institution_normalize::run_institution_normalize_files(
        &config,
        &options.positional,
        run_options,
        |event| {
            if view == "developer" {
                if let Ok(encoded) = serde_json::to_string(event) {
                    eprintln!("{encoded}");
                }
            } else if view == "user" {
                let current = event
                    .progress
                    .current_file
                    .unwrap_or(event.progress.processed_files);
                eprintln!(
                    "{}（进度 {current}/{}, 成功 {}，失败 {}）",
                    event.user_message,
                    event.progress.total_files,
                    event.progress.completed_files,
                    event.progress.failed_files
                );
            }
            emitted_events.push(event.clone());
        },
    )?;
    anyhow::ensure!(
        emitted_events.len() == report.events.len(),
        "internal progress event count mismatch"
    );

    if let Some(report_output) = report_output.as_deref() {
        write_json_report_create_new(report_output, &report)?;
    }
    let output = if view == "developer" {
        serde_json::to_value(&report)?
    } else {
        institution_normalize_user_view(&report)
    };
    let mut output = output;
    if view == "off" {
        if let Some(object) = output.as_object_mut() {
            object.remove("events");
        }
    }
    if let Some(report_output) = report_output {
        output["history_report"] = serde_json::Value::String(report_output.display().to_string());
    }
    println!("{}", serde_json::to_string_pretty(&output)?);
    Ok(())
}

fn institution_normalize_user_view(
    report: &pytools_core::institution_normalize::InstitutionNormalizeRunReport,
) -> serde_json::Value {
    serde_json::json!({
        "schema_version": report.schema_version,
        "product": report.product,
        "feature_id": report.feature_id,
        "run_id": report.run_id,
        "started_at": report.started_at,
        "status": user_friendly_normalize_run_status(&report.status),
        "mode": if report.mode == "apply" { "执行修改模式" } else { "仅生成计划，未修改文件" },
        "failure": report.failure.as_ref().map(|failure| serde_json::json!({
            "stage": user_friendly_normalize_error_stage(&failure.stage),
            "message": failure.user_message
        })),
        "progress": report.events.last().map(|event| &event.progress),
        "summary": {
            "total_inputs": report.total_inputs,
            "completed_files": report.completed_files,
            "failed_files": report.failed_files,
            "modified_files": report.modified_files,
            "branch_sheets": report.branch_sheets,
            "mapped_sheets": report.mapped_sheets,
            "mapped_rows": report.mapped_rows,
            "mapping_empty": report.mapping_empty,
            "configuration_elapsed_ms": report.configuration_elapsed_ms,
            "elapsed_ms": report.elapsed_ms
        },
        "events": report.events.iter().map(|event| serde_json::json!({
            "sequence": event.sequence,
            "timestamp": event.timestamp,
            "phase": user_friendly_normalize_phase(&event.phase),
            "status": user_friendly_normalize_event_status(&event.status),
            "message": event.user_message,
            "progress": event.progress,
            "file_name": event.file_name,
            "elapsed_ms": event.elapsed_ms
        })).collect::<Vec<_>>(),
        "files": report.files.iter().map(|file| serde_json::json!({
            "file_name": file.file_name,
            "status": user_friendly_normalize_status(&file.status),
            "branch_sheets": file.branch_sheets,
            "mapped_sheets": file.mapped_sheets,
            "mapped_rows": file.mapped_rows,
            "operation_count": file.operation_count,
            "plan_saved": file.plan_path.is_some(),
            "stages": [
                {"name": "文件检查", "elapsed_ms": file.timings.source_fingerprint_ms},
                {"name": "工作簿结构检查", "elapsed_ms": file.timings.profile_inspection_ms},
                {"name": "业务规则与修改计划", "elapsed_ms": file.timings.plan_build_ms},
                {"name": "计划保存", "elapsed_ms": file.timings.plan_persistence_ms},
                {"name": "输入一致性复核", "elapsed_ms": file.timings.source_stability_check_ms},
                {"name": "修改引擎执行", "elapsed_ms": file.timings.engine_execution_ms},
                {"name": "安全提交", "elapsed_ms": file.timings.safe_commit_ms},
                {"name": "文件总耗时", "elapsed_ms": file.timings.total_ms}
            ],
            "error_stage": file.error_stage.as_deref().map(user_friendly_normalize_error_stage),
            "message": file.user_message,
            "cleanup_warning": file.cleanup_warning
        })).collect::<Vec<_>>()
    })
}

fn user_friendly_normalize_phase(phase: &str) -> &str {
    match phase {
        "source_fingerprint" => "文件安全检查",
        "profile_inspection" => "工作簿结构检查",
        "plan_build" => "规则检查与修改计划",
        "plan_persistence" => "保存修改计划",
        "source_stability_check" => "输入一致性复核",
        "engine_execution" => "应用修改",
        "safe_commit" => "验证并提交结果",
        "configuration_read" => "读取映射配置",
        "normalization" => "机构名称规范化",
        "file_processing" => "处理工作簿",
        "task" => "任务汇总",
        _ => "处理步骤",
    }
}

fn user_friendly_normalize_run_status(status: &str) -> &str {
    match status {
        "completed" => "已完成",
        "partial_failure" => "部分文件失败",
        "failed" => "失败",
        _ => "处理中",
    }
}

fn user_friendly_normalize_event_status(status: &str) -> &str {
    match status {
        "started" => "进行中",
        "completed" => "已完成",
        "partial_failure" => "部分文件失败",
        "failed" => "失败",
        "skipped" => "已跳过",
        _ => "处理中",
    }
}

fn user_friendly_normalize_status(status: &str) -> &str {
    match status {
        "planned" => "已生成计划，尚未修改",
        "applied" => "已修改",
        "no_changes" => "无需修改",
        "failed" => "失败",
        other => other,
    }
}

fn user_friendly_normalize_error_stage(stage: &str) -> &str {
    match stage {
        "configuration_read" => "读取映射配置",
        "format_check" => "文件格式检查",
        "input_check" => "输入文件检查",
        "source_fingerprint" | "source_stability_check" => "源文件安全检查",
        "profile_inspection" => "工作簿结构检查",
        "plan_build" => "规则计算与修改计划",
        "plan_persistence" => "修改计划保存",
        "engine_execution" => "工作簿修改",
        "safe_commit" => "安全提交",
        other => other,
    }
}

fn write_json_report_create_new<T: serde::Serialize>(path: &Path, report: &T) -> Result<()> {
    if let Some(parent) = path
        .parent()
        .filter(|parent| !parent.as_os_str().is_empty())
    {
        fs::create_dir_all(parent)
            .with_context(|| format!("create report directory {}", parent.display()))?;
    }
    let mut output = OpenOptions::new()
        .write(true)
        .create_new(true)
        .open(path)
        .with_context(|| format!("create run report {} without overwriting", path.display()))?;
    output
        .write_all(&serde_json::to_vec_pretty(report)?)
        .with_context(|| format!("write run report {}", path.display()))?;
    output
        .sync_all()
        .with_context(|| format!("sync run report {}", path.display()))?;
    Ok(())
}

fn run_fx_header_batch(arguments: Vec<String>) -> Result<()> {
    let (commit_target, arguments) = extract_boolean_flag(arguments, "commit-target")?;
    let options = parse_cli_arguments(arguments)?;
    anyhow::ensure!(
        !options.positional.is_empty(),
        "provide at least one FX header workbook"
    );
    let capabilities = options.values.get("capabilities");
    let platform = options.values.get("platform");
    anyhow::ensure!(
        capabilities.is_some() == platform.is_some(),
        "--capabilities and --platform must be provided together"
    );
    anyhow::ensure!(
        !capabilities.is_some() || commit_target,
        "FX header execution requires --commit-target with matching capability evidence"
    );
    anyhow::ensure!(
        !commit_target || capabilities.is_some(),
        "--commit-target requires --capabilities and --platform"
    );
    let view = options
        .values
        .get("view")
        .map(String::as_str)
        .unwrap_or("user");
    anyhow::ensure!(
        matches!(view, "user" | "developer" | "off"),
        "--view must be user, developer, or off"
    );
    anyhow::ensure!(
        view != "off" || !options.values.contains_key("report-output"),
        "--report-output cannot be used with --view off"
    );

    let output_dir = options
        .values
        .get("output-dir")
        .map(PathBuf::from)
        .unwrap_or_else(|| PathBuf::from("output"));
    let log_dir = options
        .values
        .get("log-dir")
        .map(PathBuf::from)
        .unwrap_or_else(|| PathBuf::from("logs"));
    let plan_dir = options
        .values
        .get("plan-dir")
        .map(PathBuf::from)
        .unwrap_or_else(|| unique_menu_action_dir(&output_dir.join("plans").join("PT-7-4"), "run"));
    fs::create_dir_all(&output_dir)
        .with_context(|| format!("create output directory {}", output_dir.display()))?;
    fs::create_dir_all(&plan_dir)
        .with_context(|| format!("create plan directory {}", plan_dir.display()))?;

    let run_started = Instant::now();
    let started_at = Local::now().to_rfc3339();
    let nonce = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map(|duration| duration.as_nanos())
        .unwrap_or_default();
    let run_id = format!("pytools-pt-7-4-{}-{nonce}", std::process::id());
    let report_output = if view == "off" {
        None
    } else {
        let path = options
            .values
            .get("report-output")
            .map(PathBuf::from)
            .unwrap_or_else(|| {
                log_dir.join(format!(
                    "PT-7-4_{}_{}_{}.json",
                    Local::now().format("%Y%m%dT%H%M%S"),
                    std::process::id(),
                    nonce
                ))
            });
        anyhow::ensure!(
            path.extension()
                .and_then(|extension| extension.to_str())
                .is_some_and(|extension| extension.eq_ignore_ascii_case("json")),
            "--report-output must use a .json extension"
        );
        anyhow::ensure!(
            !path.exists(),
            "run history file already exists; refusing to overwrite {}",
            path.display()
        );
        Some(path)
    };

    let mut events = Vec::<SummaryRunEvent>::new();
    let mut progress = SummaryProgress {
        total_inputs: options.positional.len(),
        total_files: options.positional.len(),
        ..SummaryProgress::default()
    };
    let mut items = Vec::with_capacity(options.positional.len());
    let mut completed_files = 0usize;
    let mut failed_files = 0usize;
    let mut planned_files = 0usize;
    let mut committed_files = 0usize;
    let mut fx_sheets = 0usize;
    let mut modified_sheets = 0usize;

    for (index, input) in options.positional.iter().enumerate() {
        progress.current_file = Some(index + 1);
        progress.processed_files = index;
        progress.completed_files = completed_files;
        progress.failed_files = failed_files;
        if view == "user" {
            eprintln!(
                "[处理中] {}/{} 检查外汇页眉",
                index + 1,
                options.positional.len()
            );
        }
        emit_fx_header_event(
            &mut events,
            view,
            run_started,
            "input_validation",
            "started",
            format!(
                "开始处理第 {}/{} 个工作簿",
                index + 1,
                options.positional.len()
            ),
            &progress,
            None,
            &["pytools-core"],
            None,
            None,
        );

        let file_started = Instant::now();
        let mut current_stage = "input_validation";
        let item_result = (|| -> Result<serde_json::Value> {
            anyhow::ensure!(
                input
                    .extension()
                    .is_some_and(|extension| extension.eq_ignore_ascii_case("xlsx")),
                "PT-7-4 currently accepts .xlsx only"
            );
            let metadata = fs::symlink_metadata(input)
                .with_context(|| format!("inspect input workbook {}", input.display()))?;
            anyhow::ensure!(
                metadata.is_file() && !metadata.file_type().is_symlink(),
                "input workbook must be a regular, non-symlink file: {}",
                input.display()
            );
            let expected_target_hash = if commit_target {
                Some(sha256_file(input)?)
            } else {
                None
            };
            let candidate = unique_workbook_candidate(input, &output_dir, "fx_header");
            current_stage = "profile_inspection";
            let profile_started = Instant::now();
            let profile = excel_editor::inspect_xlsx_profile(input)?;
            emit_fx_header_event(
                &mut events,
                view,
                run_started,
                "profile_inspection",
                "completed",
                "已检查工作簿格式与对象特征",
                &progress,
                Some(profile_started),
                &["excel-editor", "excel-reader"],
                Some("inspect_xlsx_profile"),
                None,
            );

            current_stage = "plan_build";
            let plan_started = Instant::now();
            let inspection = pytools_core::fx_header::build_fx_header_plan(
                input,
                &candidate,
                profile,
                format!("{run_id}-{:04}", index + 1),
            )?;
            if let Some(expected_hash) = expected_target_hash.as_deref() {
                verify_planning_snapshot(input, Some(expected_hash))?;
            }
            emit_fx_header_event(
                &mut events,
                view,
                run_started,
                "plan_build",
                if inspection.plan.is_some() {
                    "completed"
                } else {
                    "no_changes"
                },
                if inspection.plan.is_some() {
                    "已完成页眉检查并生成修改计划"
                } else {
                    "没有需要修改的外汇页眉"
                },
                &progress,
                Some(plan_started),
                &["pytools-core", "excel-reader"],
                Some("read_header_inventory"),
                None,
            );

            let plan_file = if let Some(plan) = inspection.plan.as_ref() {
                current_stage = "plan_persistence";
                let path = plan_dir.join(format!("target-{:04}.plan.json", index + 1));
                let persistence_started = Instant::now();
                write_json_report_create_new(&path, plan)?;
                emit_fx_header_event(
                    &mut events,
                    view,
                    run_started,
                    "plan_persistence",
                    "completed",
                    "已保存页眉修改计划",
                    &progress,
                    Some(persistence_started),
                    &["pytools-core"],
                    None,
                    None,
                );
                Some(path)
            } else {
                None
            };

            let execution = if let Some(plan) = inspection.plan.as_ref() {
                current_stage = "engine_execution";
                let execution_started = Instant::now();
                let execution = execute_optional_plan(&options, Some(plan))?;
                let engine = execution.get("engine").and_then(serde_json::Value::as_str);
                emit_fx_header_event(
                    &mut events,
                    view,
                    run_started,
                    "engine_execution",
                    "completed",
                    "已通过 ExcelModifier 修改并验证候选工作簿",
                    &progress,
                    Some(execution_started),
                    &["pytools-core", "excel-editor", "ooxml-patch"],
                    engine,
                    None,
                );
                execution
            } else {
                serde_json::Value::String("no_changes".to_owned())
            };

            let commit = if commit_target && inspection.plan.is_some() {
                current_stage = "safe_commit";
                let commit_started = Instant::now();
                let commit = replace_target_from_candidate(
                    &candidate,
                    input,
                    expected_target_hash
                        .as_deref()
                        .context("target snapshot missing before commit")?,
                )?;
                emit_fx_header_event(
                    &mut events,
                    view,
                    run_started,
                    "safe_commit",
                    "completed",
                    "已验证候选并安全提交到原工作簿",
                    &progress,
                    Some(commit_started),
                    &["pytools-core", "file-commit", "excel-reader", "calamine"],
                    execution.get("engine").and_then(serde_json::Value::as_str),
                    None,
                );
                Some(commit)
            } else {
                None
            };

            Ok(serde_json::json!({
                "input": input,
                "candidate": if inspection.plan.is_some() { Some(&candidate) } else { None },
                "plan_file": plan_file,
                "status": if commit.is_some() {
                    "committed"
                } else if inspection.plan.is_none() {
                    "no_changes"
                } else {
                    "planned"
                },
                "fx_sheets": inspection.fx_sheets,
                "modified_sheets": inspection.modified_sheets,
                "execution": if inspection.plan.is_some() { execution } else { serde_json::Value::String("no_changes".to_owned()) },
                "target_commit": commit,
                "elapsed_ms": file_started.elapsed().as_millis(),
            }))
        })();

        match item_result {
            Ok(item) => {
                if item["plan_file"].is_string() {
                    planned_files += 1;
                }
                if item["status"] == "committed" {
                    committed_files += 1;
                }
                completed_files += 1;
                fx_sheets += item["fx_sheets"].as_u64().unwrap_or_default() as usize;
                modified_sheets += item["modified_sheets"].as_u64().unwrap_or_default() as usize;
                progress.processed_files = index + 1;
                progress.completed_files = completed_files;
                items.push(item);
                let item = items.last().expect("successful item was just appended");
                let (status, message) = match item["status"].as_str().unwrap_or_default() {
                    "committed" => ("completed", "当前工作簿的页眉修改已安全提交"),
                    "planned" => ("planned", "当前工作簿的页眉修改计划已保存，原文件未修改"),
                    _ => ("no_changes", "当前工作簿无需修改"),
                };
                emit_fx_header_event(
                    &mut events,
                    view,
                    run_started,
                    "file_summary",
                    status,
                    message,
                    &progress,
                    Some(file_started),
                    &["pytools-core"],
                    None,
                    None,
                );
            }
            Err(error) => {
                failed_files += 1;
                progress.processed_files = index + 1;
                progress.failed_files = failed_files;
                emit_fx_header_event(
                    &mut events,
                    view,
                    run_started,
                    current_stage,
                    "failed",
                    "当前工作簿处理失败；请查看开发者运行记录中的失败阶段",
                    &progress,
                    Some(file_started),
                    &["pytools-core"],
                    None,
                    Some("fx_header_stage_error"),
                );
                items.push(serde_json::json!({
                    "input": input,
                    "status": "failed",
                    "error_stage": current_stage,
                    "error": format!("{error:#}"),
                    "elapsed_ms": file_started.elapsed().as_millis(),
                }));
            }
        }
    }

    let total_files = options.positional.len();
    let status = if failed_files == 0 {
        "succeeded"
    } else if completed_files == 0 {
        "failed"
    } else {
        "partial"
    };
    progress.current_file = None;
    progress.processed_files = total_files;
    progress.completed_files = completed_files;
    progress.failed_files = failed_files;
    emit_fx_header_event(
        &mut events,
        view,
        run_started,
        "run_summary",
        status,
        format!("外汇页眉处理完成：成功 {completed_files}，失败 {failed_files}"),
        &progress,
        None,
        &["pytools-core"],
        None,
        None,
    );

    let mut result = serde_json::json!({
        "schema_version": 1,
        "product": "pytools",
        "feature_id": "PT-7-4",
        "function": "fx-header",
        "run_id": run_id,
        "started_at": started_at,
        "status": status,
        "mode": if commit_target { "commit_target" } else { "plan_only" },
        "files": total_files,
        "completed_files": completed_files,
        "failed_files": failed_files,
        "planned_files": planned_files,
        "committed_files": committed_files,
        "saved_files": committed_files,
        "fx_sheets": fx_sheets,
        "modified_sheets": modified_sheets,
        "elapsed_ms": run_started.elapsed().as_millis(),
        "plan_dir": plan_dir,
        "items": items,
    });

    let mut history_warning = None::<String>;
    if let Some(path) = report_output.as_deref() {
        let safe_items = result["items"]
            .as_array()
            .into_iter()
            .flatten()
            .map(|item| {
                serde_json::json!({
                    "status": item["status"],
                    "error_stage": item.get("error_stage"),
                    "error_type": if item["status"] == "failed" { Some("fx_header_stage_error") } else { None },
                    "fx_sheets": item.get("fx_sheets"),
                    "modified_sheets": item.get("modified_sheets"),
                    "elapsed_ms": item.get("elapsed_ms"),
                })
            })
            .collect::<Vec<_>>();
        let history = serde_json::json!({
            "schema_version": 1,
            "product": "pytools",
            "feature_id": "PT-7-4",
            "function": "fx-header",
            "run_id": result["run_id"],
            "started_at": result["started_at"],
            "status": result["status"],
            "mode": result["mode"],
            "summary": {
                "total_files": total_files,
                "completed_files": completed_files,
                "failed_files": failed_files,
                "planned_files": planned_files,
                "committed_files": committed_files,
                "fx_sheets": fx_sheets,
                "modified_sheets": modified_sheets,
                "elapsed_ms": run_started.elapsed().as_millis(),
            },
            "items": safe_items,
            "events": events,
        });
        match write_json_report_create_new(path, &history) {
            Ok(()) => result["history_report"] = serde_json::json!(path),
            Err(error) => {
                let warning = format!("could not save PT-7-4 run history: {error:#}");
                eprintln!("pytools: warning: {warning}");
                history_warning = Some(warning);
            }
        }
    }

    let output = match view {
        "developer" => {
            result["events"] = serde_json::to_value(&events)?;
            if let Some(warning) = history_warning {
                result["history_warning"] = serde_json::json!(warning);
            }
            result
        }
        "user" | "off" => {
            let mut user = fx_header_user_view(
                &result,
                run_started.elapsed().as_millis(),
                if view == "user" { Some(&events) } else { None },
            );
            user["run_id"] = result["run_id"].clone();
            user["started_at"] = result["started_at"].clone();
            if let Some(path) = result.get("history_report") {
                user["history_report"] = path.clone();
            }
            if let Some(warning) = history_warning {
                user["history_warning"] = serde_json::json!(warning);
            }
            user
        }
        _ => unreachable!("view was validated"),
    };
    println!("{}", serde_json::to_string_pretty(&output)?);
    Ok(())
}

fn emit_fx_header_event(
    events: &mut Vec<SummaryRunEvent>,
    view: &str,
    run_started: Instant,
    phase: &str,
    status: &str,
    user_message: impl Into<String>,
    progress: &SummaryProgress,
    stage_started: Option<Instant>,
    components: &[&str],
    actual_engine: Option<&str>,
    error_type: Option<&str>,
) {
    if view == "off" {
        return;
    }
    let event = SummaryRunEvent {
        sequence: events.len() + 1,
        timestamp: Local::now().to_rfc3339(),
        phase: phase.to_owned(),
        status: status.to_owned(),
        user_message: user_message.into(),
        progress: progress.clone(),
        file_name: None,
        stage_elapsed_ms: stage_started.map(|started| started.elapsed().as_millis()),
        total_elapsed_ms: run_started.elapsed().as_millis(),
        components: components.iter().map(|value| (*value).to_owned()).collect(),
        actual_engine: actual_engine.map(str::to_owned),
        error_type: error_type.map(str::to_owned),
    };
    if view == "developer" {
        if let Ok(encoded) = serde_json::to_string(&event) {
            eprintln!("{encoded}");
        }
    } else {
        eprintln!(
            "{}（进度 {}/{}, 已完成 {}，失败 {}，已耗时 {} ms）",
            event.user_message,
            event.progress.processed_files,
            event.progress.total_files,
            event.progress.completed_files,
            event.progress.failed_files,
            event.total_elapsed_ms
        );
    }
    events.push(event);
}

fn fx_header_user_view(
    result: &serde_json::Value,
    elapsed_ms: u128,
    events: Option<&[SummaryRunEvent]>,
) -> serde_json::Value {
    let mut user = serde_json::json!({
        "schema_version": 1,
        "product": "pytools",
        "feature_id": "PT-7-4",
        "status": result["status"],
        "mode": result["mode"],
        "summary": {
            "total_files": result["files"],
            "completed_files": result["completed_files"],
            "failed_files": result["failed_files"],
            "planned_files": result["planned_files"],
            "committed_files": result["committed_files"],
            "fx_sheets": result["fx_sheets"],
            "modified_sheets": result["modified_sheets"],
            "elapsed_ms": elapsed_ms,
        }
    });
    if let Some(events) = events {
        user["events"] = serde_json::Value::Array(
            events
                .iter()
                .map(|event| {
                    serde_json::json!({
                        "sequence": event.sequence,
                        "timestamp": event.timestamp,
                        "phase": fx_header_user_phase(&event.phase),
                        "status": event.status,
                        "message": event.user_message,
                        "progress": event.progress,
                        "elapsed_ms": event.stage_elapsed_ms,
                        "total_elapsed_ms": event.total_elapsed_ms,
                    })
                })
                .collect(),
        );
    }
    user
}

fn fx_header_user_phase(phase: &str) -> &str {
    match phase {
        "input_validation" | "profile_inspection" => "输入检查",
        "plan_build" => "页眉检查",
        "plan_persistence" => "计划保存",
        "engine_execution" => "页眉修改",
        "safe_commit" => "文件提交",
        "file_summary" => "文件完成",
        "run_summary" => "任务汇总",
        other => other,
    }
}

fn run_fx_header_plan(arguments: Vec<String>) -> Result<()> {
    let options = parse_cli_arguments(arguments)?;
    let input = required_path(&options, "input")?;
    let output = required_path(&options, "output")?;
    let profile = excel_editor::inspect_xlsx_profile(&input)?;
    let report = pytools_core::fx_header::build_fx_header_plan(
        &input,
        &output,
        profile,
        "pytools-fx-header",
    )?;
    write_optional_plan(&options, &report.plan)?;
    let execution = execute_optional_plan(&options, report.plan.as_ref())?;
    println!(
        "{}",
        serde_json::to_string_pretty(&serde_json::json!({
            "fx_sheets": report.fx_sheets,
            "modified_sheets": report.modified_sheets,
            "plan_written": report.plan.is_some(),
            "execution": execution,
        }))?
    );
    Ok(())
}

fn write_optional_plan(
    options: &ParsedCli,
    plan: &Option<excel_editor::ExcelEditPlan>,
) -> Result<()> {
    if let (Some(path), Some(plan)) = (options.values.get("plan-output"), plan) {
        if let Some(parent) = Path::new(path).parent() {
            if !parent.as_os_str().is_empty() {
                fs::create_dir_all(parent)?;
            }
        }
        fs::write(path, serde_json::to_vec_pretty(plan)?)?;
    }
    Ok(())
}

fn execute_optional_plan(
    options: &ParsedCli,
    plan: Option<&excel_editor::ExcelEditPlan>,
) -> Result<serde_json::Value> {
    let capabilities = options.values.get("capabilities");
    let platform = options.values.get("platform");
    anyhow::ensure!(
        capabilities.is_some() == platform.is_some(),
        "--capabilities and --platform must be provided together"
    );
    let Some(plan) = plan else {
        return Ok(serde_json::json!("no_changes"));
    };
    match (capabilities, platform) {
        (Some(capabilities), Some(platform)) => Ok(serde_json::to_value(
            apply_plan_with_capability_file(plan, capabilities, platform)?,
        )?),
        _ => Ok(serde_json::json!("plan_only")),
    }
}

fn run_rename_sheets_plan(arguments: Vec<String>) -> Result<()> {
    let options = parse_cli_arguments(arguments)?;
    let config = required_path(&options, "config")?;
    let input = required_path(&options, "input")?;
    let output = required_path(&options, "output")?;
    let report = build_sheet_rename_plan(&config, &input, &output)?;
    write_optional_plan(&options, &report.plan)?;
    let execution = execute_optional_plan(&options, report.plan.as_ref())?;
    let mut result = serde_json::to_value(report)?;
    result["execution"] = execution;
    println!("{}", serde_json::to_string_pretty(&result)?);
    Ok(())
}

fn run_rename_sheets(arguments: Vec<String>) -> Result<()> {
    let (recursive, arguments) = extract_boolean_flag(arguments, "recursive")?;
    let options = parse_cli_arguments(arguments)?;
    let config = required_path(&options, "config")?;
    let capabilities = required_path(&options, "capabilities")?;
    let platform = options
        .values
        .get("platform")
        .context("--platform is required when executing sheet renames")?;
    let report = run_batch_sheet_rename_candidate(
        &config,
        &options.positional,
        recursive,
        &capabilities,
        platform,
    )?;
    println!("{}", serde_json::to_string_pretty(&report)?);
    Ok(())
}

fn run_submission_check_cli(arguments: Vec<String>) -> Result<()> {
    let options = parse_cli_arguments(arguments)?;
    let config = required_path(&options, "config")?;
    let output_dir = options
        .values
        .get("output-dir")
        .map(PathBuf::from)
        .unwrap_or_else(|| PathBuf::from("output"));
    if options.positional.is_empty() {
        bail!("provide at least one file or directory path");
    }
    let report = run_submission_check(&config, &options.positional, &output_dir)?;
    println!("{}", serde_json::to_string_pretty(&report)?);
    Ok(())
}

fn run_convert_batch_cli(arguments: Vec<String>, excel: bool) -> Result<()> {
    let options = parse_cli_arguments(arguments)?;
    let config = required_path(&options, "config")?;
    if options.positional.is_empty() {
        bail!("provide at least one source file path");
    }
    let engine = options
        .values
        .get("engine")
        .map(String::as_str)
        .unwrap_or("auto");
    let target = options.values.get("to").cloned().unwrap_or_else(|| {
        pytools_core::config::load_global_value(
            &config,
            if excel {
                "excel目的格式"
            } else {
                "word目的格式"
            },
            if excel { "xlsx" } else { "docx" },
        )
        .unwrap_or_else(|_| if excel { "xlsx".into() } else { "docx".into() })
    });
    let result = match (excel, options.values.get("output-dir")) {
        (true, Some(output)) => {
            convert_excel_files_to(&options.positional, Path::new(output), &target, engine)?
        }
        (false, Some(output)) => {
            convert_word_files_to(&options.positional, Path::new(output), &target, engine)?
        }
        (true, None) => {
            convert_excel_files_in_source_folders(&options.positional, &target, engine)?
        }
        (false, None) => {
            convert_word_files_in_source_folders(&options.positional, &target, engine)?
        }
    };
    println!("{}", serde_json::to_string_pretty(&result)?);
    if result.failed > 0 {
        bail!("{} conversion(s) failed", result.failed);
    }
    Ok(())
}

fn run_convert_file_cli(arguments: Vec<String>) -> Result<()> {
    let options = parse_cli_arguments(arguments)?;
    let input = required_path(&options, "input")?;
    let target = options.values.get("to").context("--to is required")?;
    let output = if let Some(value) = options.values.get("output") {
        PathBuf::from(value)
    } else {
        input.with_extension(target)
    };
    let engine = options
        .values
        .get("engine")
        .map(String::as_str)
        .unwrap_or("auto");
    let validation = convert_one(&input, &output, target, engine)?;
    println!("{validation}");
    Ok(())
}

fn run_timeline_precheck_cli(arguments: Vec<String>) -> Result<()> {
    let options = parse_cli_arguments(arguments)?;
    let config = required_path(&options, "config")?;
    let report = run_timeline_precheck(config)?;
    println!("{}", serde_json::to_string_pretty(&report)?);
    if report.errors > 0 {
        bail!("timeline precheck found {} error(s)", report.errors);
    }
    Ok(())
}

fn run_timeline_compare_cli(arguments: Vec<String>) -> Result<()> {
    let options = parse_cli_arguments(arguments)?;
    let config = required_path(&options, "config")?;
    let template = required_path(&options, "template")?;
    let global = load_global(&config)?;
    let rules = pytools_core::config::load_timeline_rules(&config)?;
    let output = run_header_path_compare(&template, &options.positional, &rules, &global)?;
    println!("{}", serde_json::json!({"output": output}));
    Ok(())
}

fn run_wide_summary_cli(arguments: Vec<String>, row_suffix: bool) -> Result<()> {
    let options = parse_cli_arguments(arguments)?;
    let config = required_path(&options, "config")?;
    let global = load_global(&config)?;
    let rules = pytools_core::config::load_timeline_rules(&config)?;
    let maps = pytools_core::config::load_path_maps(&config)?;
    let output = run_wide_summary(&options.positional, &rules, &maps, &global, row_suffix)?;
    let mut append_plans = Vec::new();
    if let Some(summary_path) = output.as_deref() {
        for (index, plan) in
            build_wide_target_append_plans(summary_path, &rules, &global.output_dir)?
                .into_iter()
                .enumerate()
        {
            let stem = summary_path
                .file_stem()
                .unwrap_or_default()
                .to_string_lossy();
            let plan_path = global
                .output_dir
                .join(format!("{stem}_target_append_{}.json", index + 1));
            fs::write(&plan_path, serde_json::to_vec_pretty(&plan)?)?;
            append_plans.push(plan_path);
        }
    }
    println!(
        "{}",
        serde_json::json!({"output": output, "target_append_plans": append_plans})
    );
    Ok(())
}

fn run_summary_used_range_cli(arguments: Vec<String>) -> Result<()> {
    let options = parse_cli_arguments(arguments)?;
    let output_dir = options
        .values
        .get("output-dir")
        .map(PathBuf::from)
        .unwrap_or_else(|| PathBuf::from("output"));
    let log_view = options
        .values
        .get("log-view")
        .map(String::as_str)
        .unwrap_or("user");
    anyhow::ensure!(
        matches!(log_view, "user" | "developer" | "off"),
        "--log-view must be user, developer, or off"
    );
    let log_dir = options
        .values
        .get("log-dir")
        .map(PathBuf::from)
        .unwrap_or_else(|| PathBuf::from("logs"));
    if options.positional.is_empty() {
        bail!("provide at least one workbook path");
    }

    let run_started = Instant::now();
    let started_at = Local::now().to_rfc3339();
    let nonce = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map(|duration| duration.as_nanos())
        .unwrap_or_default();
    let run_id = format!("pytools-pt-6-1-{}-{nonce}", std::process::id());
    let mut events = Vec::<SummaryRunEvent>::new();
    let result = if log_view == "off" {
        summarize_used_ranges(&options.positional, &output_dir)
    } else {
        summarize_used_ranges_with_progress(&options.positional, &output_dir, |event| {
            if log_view == "developer" {
                match serde_json::to_string(&event) {
                    Ok(encoded) => eprintln!("{encoded}"),
                    Err(_) => eprintln!(
                        "{{\"phase\":\"progress_event\",\"status\":\"serialization_failed\"}}"
                    ),
                }
            } else {
                eprintln!("{}", event.user_message);
            }
            events.push(event);
        })
    };

    if result.is_err()
        && log_view != "off"
        && !events.last().is_some_and(|event| event.status == "failed")
    {
        let prior = events.last();
        events.push(SummaryRunEvent {
            sequence: events.len() + 1,
            timestamp: Local::now().to_rfc3339(),
            phase: "summary".to_owned(),
            status: "failed".to_owned(),
            user_message: "按使用区域汇总失败".to_owned(),
            progress: prior
                .map(|event| event.progress.clone())
                .unwrap_or_else(|| SummaryProgress {
                    total_inputs: options.positional.len(),
                    ..SummaryProgress::default()
                }),
            file_name: prior.and_then(|event| event.file_name.clone()),
            stage_elapsed_ms: None,
            total_elapsed_ms: run_started.elapsed().as_millis(),
            components: vec!["pytools-core".to_owned()],
            actual_engine: None,
            error_type: Some("summary_error".to_owned()),
        });
        let event = events.last().expect("failure event was just appended");
        if log_view == "developer" {
            eprintln!("{}", serde_json::to_string(event)?);
        } else {
            eprintln!("{}", event.user_message);
        }
    }

    let run_log_path = if log_view == "off" {
        None
    } else {
        let path = log_dir.join(format!("PT-6-1_{run_id}.json"));
        let status = match &result {
            Ok(summary) if summary.saved => "succeeded",
            Ok(_) => "skipped",
            Err(_) => "failed",
        };
        let failure = result.as_ref().err().map(|_| {
            let event = events.last();
            serde_json::json!({
                "stage": event.map(|event| event.phase.as_str()).unwrap_or("summary"),
                "error_type": event.and_then(|event| event.error_type.as_deref()).unwrap_or("summary_error")
            })
        });
        let report = serde_json::json!({
            "schema_version": 1,
            "product": "pytools",
            "feature_id": "PT-6-1",
            "function": "summary-used-range",
            "run_id": &run_id,
            "started_at": &started_at,
            "status": status,
            "total_inputs": options.positional.len(),
            "completed_files": events.last().map(|event| event.progress.completed_files).unwrap_or(0),
            "failed_files": events.last().map(|event| event.progress.failed_files).unwrap_or(0),
            "summary": result.as_ref().ok(),
            "failure": failure,
            "elapsed_ms": run_started.elapsed().as_millis(),
            "events": events,
        });
        match write_json_report_create_new(&path, &report) {
            Ok(()) => Some(path),
            Err(error) => {
                eprintln!("pytools: warning: could not save PT-6-1 run history: {error:#}");
                None
            }
        }
    };

    match result {
        Ok(result) => {
            let mut output = serde_json::to_value(result)?;
            if log_view != "off" {
                output["run_id"] = serde_json::Value::String(run_id);
                if let Some(path) = run_log_path {
                    output["run_log_path"] = serde_json::Value::String(path.display().to_string());
                    if log_view == "user" {
                        eprintln!("运行记录已保存：{}", path.display());
                    }
                }
            }
            println!("{}", serde_json::to_string_pretty(&output)?);
            Ok(())
        }
        Err(error) => {
            if log_view == "user" {
                if let Some(path) = run_log_path {
                    eprintln!("运行记录已保存：{}", path.display());
                }
            }
            Err(error)
        }
    }
}

fn run_summary_comments_cli(arguments: Vec<String>) -> Result<()> {
    let options = parse_cli_arguments(arguments)?;
    let template = required_path(&options, "template")?;
    let output_dir = options
        .values
        .get("output-dir")
        .map(PathBuf::from)
        .unwrap_or_else(|| PathBuf::from("output"));
    let log_view = options
        .values
        .get("log-view")
        .map(String::as_str)
        .unwrap_or("user");
    anyhow::ensure!(
        matches!(log_view, "user" | "developer" | "off"),
        "--log-view must be user, developer, or off"
    );
    let log_dir = options
        .values
        .get("log-dir")
        .map(PathBuf::from)
        .unwrap_or_else(|| PathBuf::from("logs"));
    if options.positional.is_empty() {
        bail!("provide at least one source workbook path");
    }

    let run_started = Instant::now();
    let started_at = Local::now().to_rfc3339();
    let nonce = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map(|duration| duration.as_nanos())
        .unwrap_or_default();
    let run_id = format!("pytools-pt-6-2-{}-{nonce}", std::process::id());
    let mut events = Vec::<SummaryRunEvent>::new();
    let result = if log_view == "off" {
        summarize_by_comments(&template, &options.positional, &output_dir)
    } else {
        summarize_by_comments_with_progress(&template, &options.positional, &output_dir, |event| {
            if log_view == "developer" {
                match serde_json::to_string(&event) {
                    Ok(encoded) => eprintln!("{encoded}"),
                    Err(_) => eprintln!(
                        "{{\"phase\":\"progress_event\",\"status\":\"serialization_failed\"}}"
                    ),
                }
            } else {
                eprintln!("{}", event.user_message);
            }
            events.push(event);
        })
    };

    if result.is_err()
        && log_view != "off"
        && !events.last().is_some_and(|event| event.status == "failed")
    {
        let prior = events.last();
        events.push(SummaryRunEvent {
            sequence: events.len() + 1,
            timestamp: Local::now().to_rfc3339(),
            phase: "summary".to_owned(),
            status: "failed".to_owned(),
            user_message: "按批注汇总失败".to_owned(),
            progress: prior
                .map(|event| event.progress.clone())
                .unwrap_or_else(|| SummaryProgress {
                    total_inputs: options.positional.len(),
                    ..SummaryProgress::default()
                }),
            file_name: prior.and_then(|event| event.file_name.clone()),
            stage_elapsed_ms: None,
            total_elapsed_ms: run_started.elapsed().as_millis(),
            components: vec!["pytools-core".to_owned()],
            actual_engine: None,
            error_type: Some("summary_error".to_owned()),
        });
        let event = events.last().expect("failure event was just appended");
        if log_view == "developer" {
            eprintln!("{}", serde_json::to_string(event)?);
        } else {
            eprintln!("{}", event.user_message);
        }
    }

    let run_log_path = if log_view == "off" {
        None
    } else {
        let path = log_dir.join(format!("PT-6-2_{run_id}.json"));
        let status = match &result {
            Ok(summary) if summary.saved => "succeeded",
            Ok(_) => "skipped",
            Err(_) => "failed",
        };
        let failure = result.as_ref().err().map(|_| {
            let event = events.last();
            serde_json::json!({
                "stage": event.map(|event| event.phase.as_str()).unwrap_or("summary"),
                "error_type": event.and_then(|event| event.error_type.as_deref()).unwrap_or("summary_error")
            })
        });
        let report = serde_json::json!({
            "schema_version": 1,
            "product": "pytools",
            "feature_id": "PT-6-2",
            "function": "summary-comments",
            "run_id": &run_id,
            "started_at": &started_at,
            "status": status,
            "total_inputs": options.positional.len(),
            "completed_files": events.last().map(|event| event.progress.completed_files).unwrap_or(0),
            "failed_files": events.last().map(|event| event.progress.failed_files).unwrap_or(0),
            "summary": result.as_ref().ok(),
            "failure": failure,
            "elapsed_ms": run_started.elapsed().as_millis(),
            "events": events,
        });
        match write_json_report_create_new(&path, &report) {
            Ok(()) => Some(path),
            Err(error) => {
                eprintln!("pytools: warning: could not save PT-6-2 run history: {error:#}");
                None
            }
        }
    };

    match result {
        Ok(result) => {
            let mut output = serde_json::to_value(result)?;
            if log_view != "off" {
                output["run_id"] = serde_json::Value::String(run_id);
                if let Some(path) = run_log_path {
                    output["run_log_path"] = serde_json::Value::String(path.display().to_string());
                    if log_view == "user" {
                        eprintln!("运行记录已保存：{}", path.display());
                    }
                }
            }
            println!("{}", serde_json::to_string_pretty(&output)?);
            Ok(())
        }
        Err(error) => {
            if log_view == "user" {
                if let Some(path) = run_log_path {
                    eprintln!("运行记录已保存：{}", path.display());
                }
            }
            Err(error)
        }
    }
}

fn run_archive_cli(arguments: Vec<String>, operation: &str) -> Result<()> {
    let options = parse_cli_arguments(arguments)?;
    let output_dir = options
        .values
        .get("output-dir")
        .map(PathBuf::from)
        .unwrap_or_else(|| PathBuf::from("output"));
    let result = match operation {
        "date-preview" => archive_by_date_preview(
            &options.positional,
            &required_path(&options, "target-root")?,
            &output_dir,
            options.values.get("config").map(Path::new),
        )?,
        "type-preview" => archive_by_type_preview(
            &options.positional,
            &required_path(&options, "target-root")?,
            &output_dir,
            &required_path(&options, "config")?,
        )?,
        "date-copy" => archive_by_date_copy(
            &options.positional,
            &required_path(&options, "target-root")?,
            &output_dir,
            options.values.get("config").map(Path::new),
        )?,
        "type-copy" => archive_by_type_copy(
            &options.positional,
            &required_path(&options, "target-root")?,
            &output_dir,
            &required_path(&options, "config")?,
        )?,
        "plan" => execute_archive_plan(&required_path(&options, "plan")?, &output_dir)?,
        _ => unreachable!(),
    };
    println!("{}", serde_json::to_string_pretty(&result)?);
    Ok(())
}

struct ParsedCli {
    values: std::collections::BTreeMap<String, String>,
    positional: Vec<PathBuf>,
}
fn parse_cli_arguments(arguments: Vec<String>) -> Result<ParsedCli> {
    let mut values = std::collections::BTreeMap::new();
    let mut positional = Vec::new();
    let mut args = arguments.into_iter();
    while let Some(value) = args.next() {
        if value == "--help" || value == "-h" {
            print_usage();
            return Ok(ParsedCli { values, positional });
        }
        if let Some(name) = value.strip_prefix("--") {
            let argument = args
                .next()
                .with_context(|| format!("--{name} requires a value"))?;
            values.insert(name.to_owned(), argument);
        } else if value.starts_with('-') {
            bail!("unknown option {value:?}");
        } else {
            positional.push(PathBuf::from(value));
        }
    }
    Ok(ParsedCli { values, positional })
}

fn extract_boolean_flag(arguments: Vec<String>, name: &str) -> Result<(bool, Vec<String>)> {
    let flag = format!("--{name}");
    let inline_prefix = format!("{flag}=");
    anyhow::ensure!(
        !arguments
            .windows(2)
            .any(|pair| { pair[0] == flag && matches!(pair[1].as_str(), "true" | "false") }),
        "{flag} is a boolean switch and does not take a value"
    );
    let mut found = false;
    let mut filtered = Vec::with_capacity(arguments.len());
    for argument in arguments {
        if argument == flag {
            anyhow::ensure!(!found, "{flag} may only be specified once");
            found = true;
        } else if argument.starts_with(&inline_prefix) {
            bail!("{flag} is a boolean switch and does not take a value");
        } else {
            filtered.push(argument);
        }
    }
    Ok((found, filtered))
}
fn required_path(options: &ParsedCli, name: &str) -> Result<PathBuf> {
    options
        .values
        .get(name)
        .map(PathBuf::from)
        .with_context(|| format!("--{name} is required"))
}

fn unique_workbook_candidate(input: &Path, output_dir: &Path, suffix: &str) -> PathBuf {
    use std::hash::{Hash, Hasher};
    let stem = input.file_stem().unwrap_or_default().to_string_lossy();
    let mut hasher = std::collections::hash_map::DefaultHasher::new();
    input.hash(&mut hasher);
    let base = format!("{stem}_{suffix}_{:08x}", hasher.finish() as u32);
    let mut candidate = output_dir.join(format!("{base}.xlsx"));
    let mut index = 1usize;
    while candidate.exists() {
        candidate = output_dir.join(format!("{base}_{index}.xlsx"));
        index += 1;
    }
    candidate
}

fn run_local_filename_map(arguments: Vec<String>) -> Result<()> {
    let mut arguments = arguments.into_iter();
    let mut config_path = None::<PathBuf>;
    let mut files = Vec::new();
    while let Some(argument) = arguments.next() {
        match argument.as_str() {
            "--config" => {
                config_path = Some(PathBuf::from(
                    arguments.next().context("--config requires a path")?,
                ));
            }
            "--help" | "-h" => {
                print_usage();
                return Ok(());
            }
            value if value.starts_with('-') => bail!("unknown option {value:?}"),
            path => files.push(PathBuf::from(path)),
        }
    }
    let config_path = config_path.context("--config is required")?;
    if files.is_empty() {
        bail!("provide at least one file path");
    }

    let result = run_local_filename_map_candidate(&config_path, &files)?;
    println!("{}", serde_json::to_string_pretty(&result)?);
    Ok(())
}

fn run_batch_rename(arguments: Vec<String>) -> Result<()> {
    let mut arguments = arguments.into_iter();
    let mut config_path = None::<PathBuf>;
    let mut log_dir = None::<PathBuf>;
    let mut paths = Vec::new();
    while let Some(argument) = arguments.next() {
        match argument.as_str() {
            "--config" => {
                config_path = Some(PathBuf::from(
                    arguments.next().context("--config requires a path")?,
                ));
            }
            "--log-dir" => {
                log_dir = Some(PathBuf::from(
                    arguments.next().context("--log-dir requires a path")?,
                ));
            }
            "--help" | "-h" => {
                print_usage();
                return Ok(());
            }
            value if value.starts_with('-') => bail!("unknown option {value:?}"),
            path => paths.push(PathBuf::from(path)),
        }
    }

    let config_path = config_path.context("--config is required")?;
    let log_dir = log_dir.unwrap_or_else(|| PathBuf::from("logs"));
    let (log_path, mut log_file) = open_batch_rename_log(&log_dir)?;
    let result = run_batch_rename_business(&config_path, &paths)?;
    let log_warning = write_batch_rename_log(&mut log_file, &result.rows).err();
    let mut output = serde_json::to_value(result)?;
    if let Some(object) = output.as_object_mut() {
        object.insert(
            "log_path".to_owned(),
            serde_json::json!(log_path.to_string_lossy()),
        );
        if let Some(warning) = log_warning {
            object.insert(
                "log_warning".to_owned(),
                serde_json::json!(warning.to_string()),
            );
        }
    }
    println!("{}", serde_json::to_string_pretty(&output)?);
    Ok(())
}

fn open_batch_rename_log(log_dir: &Path) -> Result<(PathBuf, File)> {
    fs::create_dir_all(log_dir)
        .with_context(|| format!("create batch-rename log directory {}", log_dir.display()))?;
    let timestamp = Local::now().format("%Y%m%d_%H%M%S");
    let base = format!("3_3_batch_rename_{timestamp}");
    for index in 0usize.. {
        let name = if index == 0 {
            format!("{base}.log")
        } else {
            format!("{base}_{index}.log")
        };
        let path = log_dir.join(name);
        match OpenOptions::new().write(true).create_new(true).open(&path) {
            Ok(file) => return Ok((path, file)),
            Err(error) if error.kind() == std::io::ErrorKind::AlreadyExists => continue,
            Err(error) => {
                return Err(error).with_context(|| format!("create log {}", path.display()))
            }
        }
    }
    unreachable!("usize log suffix range exhausted")
}

fn write_batch_rename_log(
    file: &mut File,
    rows: &[pytools_core::batch_rename::BatchRenameRow],
) -> Result<()> {
    for row in rows {
        let (level, message) = if row.status == "renamed" {
            (
                "INFO",
                format!("ok old={} new={}", row.original_name, row.new_name),
            )
        } else {
            let file = if row.original_name.is_empty() {
                row.source_path.as_str()
            } else {
                row.original_name.as_str()
            };
            let level = if row.reason.starts_with("rename_failed") {
                "WARNING"
            } else {
                "INFO"
            };
            (level, format!("skip file={file} reason={}", row.reason))
        };
        let line = format!(
            "{} [{level}] {message}",
            Local::now().format("%Y-%m-%d %H:%M:%S")
        );
        writeln!(file, "{line}")?;
        eprintln!("{line}");
    }
    file.flush()?;
    Ok(())
}

fn run_content_prefix_rename_cli(arguments: Vec<String>) -> Result<()> {
    let Some((config_path, files)) = parse_config_and_files(arguments)? else {
        return Ok(());
    };
    let result = run_content_prefix_rename_xlsx_candidate(&config_path, &files)?;
    println!("{}", serde_json::to_string_pretty(&result)?);
    Ok(())
}

fn run_content_prefix_unrename_cli(arguments: Vec<String>) -> Result<()> {
    let Some((config_path, files)) = parse_config_and_files(arguments)? else {
        return Ok(());
    };
    let result = run_content_prefix_unrename_xlsx_candidate(&config_path, &files)?;
    println!("{}", serde_json::to_string_pretty(&result)?);
    Ok(())
}

fn parse_config_and_files(arguments: Vec<String>) -> Result<Option<(PathBuf, Vec<PathBuf>)>> {
    let mut arguments = arguments.into_iter();
    let mut config_path = None::<PathBuf>;
    let mut files = Vec::new();
    while let Some(argument) = arguments.next() {
        match argument.as_str() {
            "--config" => {
                config_path = Some(PathBuf::from(
                    arguments.next().context("--config requires a path")?,
                ));
            }
            "--help" | "-h" => {
                print_usage();
                return Ok(None);
            }
            value if value.starts_with('-') => bail!("unknown option {value:?}"),
            path => files.push(PathBuf::from(path)),
        }
    }

    let config_path = config_path.context("--config is required")?;
    if files.is_empty() {
        bail!("provide at least one file path");
    }
    Ok(Some((config_path, files)))
}

fn run_file_sort_prefix(arguments: Vec<String>) -> Result<()> {
    let mut arguments = arguments.into_iter();
    let mut config_path = None::<PathBuf>;
    let mut output_dir = None::<PathBuf>;
    let mut log_dir = None::<PathBuf>;
    let mut files = Vec::new();
    while let Some(argument) = arguments.next() {
        match argument.as_str() {
            "--config" => {
                config_path = Some(PathBuf::from(
                    arguments.next().context("--config requires a path")?,
                ));
            }
            "--output-dir" => {
                output_dir = Some(PathBuf::from(
                    arguments.next().context("--output-dir requires a path")?,
                ));
            }
            "--log-dir" => {
                log_dir = Some(PathBuf::from(
                    arguments.next().context("--log-dir requires a path")?,
                ));
            }
            "--help" | "-h" => {
                print_usage();
                return Ok(());
            }
            value if value.starts_with('-') => bail!("unknown option {value:?}"),
            path => files.push(PathBuf::from(path)),
        }
    }

    let config_path = config_path.context("--config is required")?;
    if files.is_empty() {
        bail!("provide at least one file path");
    }
    let output_dir = output_dir.unwrap_or_else(|| PathBuf::from("output"));
    let log_dir = log_dir.unwrap_or_else(|| PathBuf::from("logs"));
    let (log_path, mut log_file) = open_file_sort_log(&log_dir)?;

    let result = run_file_sort_prefix_business(&config_path, &files, &output_dir)?;
    let log_warning = write_file_sort_log(&mut log_file, &result.rows).err();
    let mut output = serde_json::to_value(result)?;
    output["log_path"] = serde_json::json!(log_path);
    if let Some(error) = log_warning {
        let warning = format!("could not write complete operation log: {error:#}");
        eprintln!("pytools: warning: {warning}");
        output["log_warning"] = serde_json::json!(warning);
    }
    println!("{}", serde_json::to_string_pretty(&output)?);
    Ok(())
}

fn open_file_sort_log(log_dir: &Path) -> Result<(PathBuf, File)> {
    fs::create_dir_all(log_dir)
        .with_context(|| format!("create log directory {}", log_dir.display()))?;
    let timestamp = Local::now().format("%Y%m%d_%H%M%S");
    let path = log_dir.join(format!("4_9_file_sort_prefix_{timestamp}.log"));
    let file = OpenOptions::new()
        .create(true)
        .append(true)
        .open(&path)
        .with_context(|| format!("open operation log {}", path.display()))?;
    Ok((path, file))
}

fn write_file_sort_log(
    file: &mut File,
    rows: &[pytools_core::file_sort_prefix::FileSortPrefixRow],
) -> Result<()> {
    for row in rows {
        let message = if let Some(config_row) = row.config_row {
            Some((
                "INFO",
                format!(
                    "ok old={} new={} row={} prefix={}",
                    row.original_name, row.new_name, config_row, row.sort_prefix
                ),
            ))
        } else if row.status == "失败" {
            Some((
                "WARNING",
                format!("failed file={} err={}", row.source_path, row.remark),
            ))
        } else {
            None
        };
        if let Some((level, message)) = message {
            let line = format!(
                "{} [{level}] {message}",
                Local::now().format("%Y-%m-%d %H:%M:%S")
            );
            writeln!(file, "{line}")?;
            eprintln!("{line}");
        }
    }
    file.flush()?;
    Ok(())
}

fn run_remove_file_sort_prefix(arguments: Vec<String>) -> Result<()> {
    let mut arguments = arguments.into_iter();
    let mut config_path = None::<PathBuf>;
    let mut output_dir = None::<PathBuf>;
    let mut log_dir = None::<PathBuf>;
    let mut files = Vec::new();
    while let Some(argument) = arguments.next() {
        match argument.as_str() {
            "--config" => {
                config_path = Some(PathBuf::from(
                    arguments.next().context("--config requires a path")?,
                ));
            }
            "--output-dir" => {
                output_dir = Some(PathBuf::from(
                    arguments.next().context("--output-dir requires a path")?,
                ));
            }
            "--log-dir" => {
                log_dir = Some(PathBuf::from(
                    arguments.next().context("--log-dir requires a path")?,
                ));
            }
            "--help" | "-h" => {
                print_usage();
                return Ok(());
            }
            value if value.starts_with('-') => bail!("unknown option {value:?}"),
            path => files.push(PathBuf::from(path)),
        }
    }

    let config_path = config_path.context("--config is required")?;
    let output_dir = output_dir.unwrap_or_else(|| PathBuf::from("output"));
    let log_dir = log_dir.unwrap_or_else(|| PathBuf::from("logs"));
    let (log_path, mut log_file) = open_file_sort_removal_log(&log_dir)?;

    let result = run_remove_file_sort_prefix_business(&config_path, &files, &output_dir)?;
    let log_warning = write_file_sort_removal_log(&mut log_file, &result.rows).err();
    let mut output = serde_json::to_value(result)?;
    output["log_path"] = serde_json::json!(log_path);
    if let Some(error) = log_warning {
        let warning = format!("could not write complete operation log: {error:#}");
        eprintln!("pytools: warning: {warning}");
        output["log_warning"] = serde_json::json!(warning);
    }
    println!("{}", serde_json::to_string_pretty(&output)?);
    Ok(())
}

fn open_file_sort_removal_log(log_dir: &Path) -> Result<(PathBuf, File)> {
    fs::create_dir_all(log_dir)
        .with_context(|| format!("create log directory {}", log_dir.display()))?;
    let timestamp = Local::now().format("%Y%m%d_%H%M%S");
    let path = log_dir.join(format!("4_10_remove_file_sort_prefix_{timestamp}.log"));
    let file = OpenOptions::new()
        .create(true)
        .append(true)
        .open(&path)
        .with_context(|| format!("open operation log {}", path.display()))?;
    Ok((path, file))
}

fn write_file_sort_removal_log(
    file: &mut File,
    rows: &[pytools_core::file_sort_prefix::FileSortPrefixRemovalRow],
) -> Result<()> {
    for row in rows {
        let message = if !row.new_name.is_empty() {
            Some((
                "INFO",
                format!("ok old={} new={}", row.original_name, row.new_name),
            ))
        } else if row.status == "失败" {
            Some((
                "WARNING",
                format!("failed file={} err={}", row.source_path, row.remark),
            ))
        } else {
            None
        };
        if let Some((level, message)) = message {
            let line = format!(
                "{} [{level}] {message}",
                Local::now().format("%Y-%m-%d %H:%M:%S")
            );
            writeln!(file, "{line}")?;
            eprintln!("{line}");
        }
    }
    file.flush()?;
    Ok(())
}

fn run_extract_sheets(arguments: Vec<String>) -> Result<()> {
    let mut arguments = arguments.into_iter();
    let mut config_path = None::<PathBuf>;
    let mut output_dir = None::<PathBuf>;
    let mut capability_path = None::<PathBuf>;
    let mut platform = None::<String>;
    let mut source_workbooks = Vec::new();
    while let Some(argument) = arguments.next() {
        match argument.as_str() {
            "--config" => {
                config_path = Some(PathBuf::from(
                    arguments.next().context("--config requires a path")?,
                ));
            }
            "--output-dir" => {
                output_dir = Some(PathBuf::from(
                    arguments.next().context("--output-dir requires a path")?,
                ));
            }
            "--capabilities" => {
                capability_path = Some(PathBuf::from(
                    arguments.next().context("--capabilities requires a path")?,
                ));
            }
            "--platform" => {
                platform = Some(arguments.next().context("--platform requires an id")?);
            }
            "--help" | "-h" => {
                print_usage();
                return Ok(());
            }
            value if value.starts_with('-') => bail!("unknown option {value:?}"),
            path => source_workbooks.push(PathBuf::from(path)),
        }
    }

    let config_path = config_path.context("--config is required")?;
    if source_workbooks.is_empty() {
        bail!("provide at least one source workbook path");
    }
    anyhow::ensure!(
        capability_path.is_some() == platform.is_some(),
        "--capabilities and --platform must be provided together"
    );
    let tasks = load_extract_tasks(&config_path).with_context(|| {
        format!(
            "load worksheet-extraction tasks from {}",
            config_path.display()
        )
    })?;
    let output_dir = match output_dir {
        Some(path) => path,
        None => {
            load_global(&config_path)
                .with_context(|| format!("load output directory from {}", config_path.display()))?
                .output_dir
        }
    };
    let result = run_extract_sheets_candidate(
        &source_workbooks,
        &tasks,
        &output_dir,
        false,
        None,
        capability_path.as_deref(),
        platform.as_deref(),
    )?;
    println!("{}", serde_json::to_string_pretty(&result)?);
    Ok(())
}

fn run_region_sum(arguments: Vec<String>) -> Result<()> {
    let mut arguments = arguments.into_iter();
    let mut config_path = None::<PathBuf>;
    let mut output_dir = None::<PathBuf>;
    let mut log_dir = None::<PathBuf>;
    let mut source_workbooks = Vec::new();
    while let Some(argument) = arguments.next() {
        match argument.as_str() {
            "--config" => {
                config_path = Some(PathBuf::from(
                    arguments.next().context("--config requires a path")?,
                ));
            }
            "--output-dir" => {
                output_dir = Some(PathBuf::from(
                    arguments.next().context("--output-dir requires a path")?,
                ));
            }
            "--log-dir" => {
                log_dir = Some(PathBuf::from(
                    arguments.next().context("--log-dir requires a path")?,
                ));
            }
            "--help" | "-h" => {
                print_usage();
                return Ok(());
            }
            value if value.starts_with('-') => bail!("unknown option {value:?}"),
            path => source_workbooks.push(PathBuf::from(path)),
        }
    }

    if source_workbooks.is_empty() {
        bail!("provide at least one source workbook path");
    }

    let global = match config_path.as_ref() {
        Some(path) => Some(
            load_global(path)
                .with_context(|| format!("load global configuration from {}", path.display()))?,
        ),
        None => None,
    };
    let output_dir = output_dir
        .or_else(|| global.as_ref().map(|config| config.output_dir.clone()))
        .unwrap_or_else(|| PathBuf::from("output"));
    let log_dir = log_dir
        .or_else(|| global.as_ref().map(|config| config.log_dir.clone()))
        .unwrap_or_else(|| PathBuf::from("logs"));

    let result = run_region_sum_check(&source_workbooks, &output_dir, &log_dir)?;
    println!(
        "{}",
        serde_json::to_string_pretty(&serde_json::json!({
            "files": result.files,
            "sheets": result.sheets,
            "errors": result.errors,
            "failed_files": result.failed_files,
            "path": result.path.map(|path| path.to_string_lossy().into_owned()),
            "log_path": result.log_path.to_string_lossy(),
        }))?
    );
    Ok(())
}

fn print_usage() {
    eprintln!(
        "Usage:\n\
         pytools [menu] [--config <config.xlsx>] [--log-view user|developer|off] [--capabilities <json> --platform <id>]\n\
         no command opens the migrated menu; available categories: 1 (timeline 1/3/4/5), 2 (dedup 1-6), 3 (print 1-4), 4 (conversion and file naming 1-10), 6 (summary 1/2), and 7 (batch 7-1/7-5)\n\
         timeline-precheck --config <workbook.xlsx>\n\
         timeline-compare --config <workbook.xlsx> --template <workbook.xlsx> <source.xlsx>...\n\
         wide-summary --config <workbook.xlsx> <source.xlsx>...\n\
         wide-summary-row-suffix --config <workbook.xlsx> <source.xlsx>...\n\
         wide-summary target writes produce append-plan JSON; execute with dedup-append-execute and matching capability evidence\n\
         summary-used-range [--output-dir <dir>] [--log-dir <dir>] [--log-view user|developer|off] <workbook>...\n\
         summary-comments --template <workbook.xlsx> [--output-dir <dir>] [--log-dir <dir>] [--log-view user|developer|off] <source.xlsx>...\n\
         dedup-comment-plan --input <workbook.xlsx> --output <candidate.xlsx> [--action check|mark|delete] [--plan-output <plan.json>] [--capabilities <json> --platform <id>]\n\
         dedup-config-plan --config <workbook.xlsx> [--output-dir <dir>] [--action check|mark|delete] [--capabilities <json> --platform <id>]\n\
         dedup-precheck --config <workbook.xlsx>\n\
         dedup-append-plan --config <workbook.xlsx> [--output-dir <dir>]\n\
         dedup-append-execute --plan <append-plan.json> [--capabilities <json> --platform <id>] [--commit-target]\n\
         print-comment-plan --input <workbook.xlsx> --output <candidate.xlsx> [--drop-comments true] [--hide-zero true] [--plan-output <plan.json>] [--capabilities <json> --platform <id>]\n\
         print-config-precheck --config <workbook.xlsx>\n\
         print-config-plan --config <workbook.xlsx> [--output-dir <plan-report-dir>] [--capabilities <json> --platform <id>]\n\
         apply-edit-plan --plan <plan.json> --capabilities <json> --platform <id>\n\
         config-init --path <new-config.xlsx>\n\
         config-repair-plan --input <config.xlsx> --output <candidate.xlsx> [--plan-output <plan.json>] [--capabilities <json> --platform <id>]\n\
         config-view --config <config.xlsx>\n\
         village-split-plan --current <xlsx> --previous <xlsx> --input <target.xlsx> --output <candidate.xlsx> [--plan-output <json>] [--capabilities <json> --platform <id>]\n\
         village-split --current <xlsx> --previous <xlsx> [--config <config.xlsx>] [--output-dir <directory>] [--log-dir <directory>] [--log-view user|developer|off] [--capabilities <json> --platform <id> [--commit-target]] <target.xlsx>...\n\
         remove-foreign-bank-plan --config <config.xlsx> --input <xlsx> --output <candidate.xlsx> [--plan-output <json>] [--capabilities <json> --platform <id>]\n\
         remove-foreign-bank --config <config.xlsx> [--plan-dir <dir>] [--view user|developer|off] [--log-dir <dir>] [--report-output <json>] [--capabilities <json> --platform <id>] <source.xlsx>...\n\
         rural-loan-plan --current <xlsx> --previous <xlsx> --output <candidate.xlsx> [--plan-output <json>] [--capabilities <json> --platform <id>]\n\
         quarterly-rural-loan-plan --current <xlsx> --previous <xlsx> --output <candidate.xlsx> [--plan-output <json>] [--capabilities <json> --platform <id>]\n\
         institution-normalize-plan --config <config.xlsx> --input <xlsx> --output <candidate.xlsx> [--plan-output <json>] [--capabilities <json> --platform <id>]\n\
         institution-normalize --config <config.xlsx> --plan-dir <dir> [--view user|developer|off] [--report-output <json>] <workbook.xlsx>...\n\
         institution-normalize --config <config.xlsx> --apply --capabilities <json> --platform <id> [--view user|developer|off] [--report-output <json>] <workbook.xlsx>...\n\
         fx-header <xlsx>... [--output-dir <dir>] [--log-dir <dir>] [--plan-dir <dir>] [--view user|developer|off] [--commit-target --capabilities <json> --platform <id>]\n\
         fx-header-plan --input <xlsx> --output <candidate.xlsx> [--plan-output <json>] [--capabilities <json> --platform <id>]\n\
         rename-sheets-plan --config <config.xlsx> --input <xlsx> --output <candidate.xlsx> [--plan-output <json>] [--capabilities <json> --platform <id>]\n\
         rename-sheets --config <config.xlsx> --capabilities <json> --platform <id> [--recursive] <file-or-directory>...\n\
         submission-check --config <config.xlsx> [--output-dir <dir>] <file-or-directory>...\n\
         archive-date-preview --target-root <dir> [--config <config.xlsx>] [--output-dir <dir>] <file-or-directory>...\n\
         archive-type-preview --config <config.xlsx> --target-root <dir> [--output-dir <dir>] <file-or-directory>...\n\
         archive-date-copy --target-root <dir> [--config <config.xlsx>] [--output-dir <dir>] <file-or-directory>...\n\
         archive-type-copy --config <config.xlsx> --target-root <dir> [--output-dir <dir>] <file-or-directory>...\n\
         archive-execute-plan --plan <archive-plan.xlsx> [--output-dir <dir>]\n\
         convert-excel --config <config.xlsx> [--to xlsx] [--engine auto|lo|native] [--output-dir <dir>] <file>...\n\
         convert-word --config <config.xlsx> [--to docx] [--engine auto|lo] [--output-dir <dir>] <file>...\n\
         convert-file --input <file> --to <extension> [--output <file>] [--engine auto|lo|native]\n\
         native = 按值重建（C01，丢弃样式/公式缓存）；Excel 公式文本保留但不重算，CSV 按文本写入\n\
         pytools extract-sheets --config <workbook.xlsx> [--output-dir <dir>] [--capabilities <json> --platform <id>] <source.xlsx>...\n\
         pytools region-sum-check [--config <workbook.xlsx>] [--output-dir <dir>] [--log-dir <dir>] <source.xlsx>...\n\
         pytools file-sort-prefix --config <workbook.xlsx> [--output-dir <dir>] [--log-dir <dir>] <file>...\n\
         pytools remove-file-sort-prefix --config <workbook.xlsx> [--output-dir <dir>] [--log-dir <dir>] [<file>...]\n\
         pytools rename-by-content --config <workbook.xlsx> <file>...\n\
         pytools unrename-by-content --config <workbook.xlsx> <file>...\n\
         pytools rename-local-map --config <workbook.xlsx> <file>...\n\
         pytools rename-batch --config <workbook.xlsx> [--log-dir <dir>] [<file-or-directory>...]\n\
         The region-sum command runs the Rust non-COM regional total check. Without --config,\n\
         output and logs default to ./output and ./logs. File-sort-prefix writes its XLSX report\n\
         under ./output unless --output-dir is supplied; its log defaults to ./logs. The\n\
         remove-file-sort-prefix command also writes under ./output and ./logs by default; it\n\
         accepts an empty file list to produce the source-compatible blank report. Content-prefix commands accept only .xlsx/.xlsm files; legacy .xls/.xlt COM reads are not part of the Rust path."
    );
}
