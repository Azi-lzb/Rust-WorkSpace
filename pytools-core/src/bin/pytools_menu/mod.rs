use anyhow::{bail, Result};
use std::{
    io::{self, Write},
    path::PathBuf,
};

struct MenuEntry {
    number: &'static str,
    label: &'static str,
    status: &'static str,
}

const MAIN_MENU: &str = "\n======== pytools 菜单 ========\n  1  时序工具\n  2  去重工具\n  3  打印工具\n  4  转换工具\n  5  配置工具\n  6  汇总工具\n  7  一二批处理工具\n  8  归档工具\n  9  快捷工具\n  0  退出\n==============================";

const TIMELINE: &[MenuEntry] = &[
    MenuEntry {
        number: "1",
        label: "表头路径比对（时序提取规则）",
        status: "Rust 入口",
    },
    MenuEntry {
        number: "2",
        label: "时序提取（时序提取规则/路径标准化映射）",
        status: "尚未实现",
    },
    MenuEntry {
        number: "3",
        label: "宽表汇总（行头加后缀）",
        status: "Rust 入口",
    },
    MenuEntry {
        number: "4",
        label: "宽表汇总（行头不加后缀）",
        status: "Rust 入口",
    },
    MenuEntry {
        number: "5",
        label: "预校验（时序提取规则/路径标准化映射）",
        status: "Rust 入口",
    },
];
const DEDUP: &[MenuEntry] = &[
    MenuEntry {
        number: "1",
        label: "按批注检查重复",
        status: "尚未实现（计划接口，未接通执行）",
    },
    MenuEntry {
        number: "2",
        label: "按批注删除重复",
        status: "尚未实现（计划接口，未接通执行）",
    },
    MenuEntry {
        number: "3",
        label: "按配置检查重复（去重追加数据配置）",
        status: "尚未实现（计划接口，未接通执行）",
    },
    MenuEntry {
        number: "4",
        label: "按配置删除重复（去重追加数据配置）",
        status: "尚未实现（计划接口，未接通执行）",
    },
    MenuEntry {
        number: "5",
        label: "按配置去重追加到目标（去重追加数据配置）",
        status: "尚未实现（等待 Excel 修改能力验收）",
    },
    MenuEntry {
        number: "6",
        label: "按配置预校验（去重追加数据配置）",
        status: "Rust 入口",
    },
];
const PRINT: &[MenuEntry] = &[
    MenuEntry {
        number: "1",
        label: "按批注打印（保留源格式，自动路由）",
        status: "尚未实现（计划接口，未接通执行）",
    },
    MenuEntry {
        number: "2",
        label: "按批注打印（快速复制，自动路由）",
        status: "尚未实现（计划接口，未接通执行）",
    },
    MenuEntry {
        number: "3",
        label: "按配置打印（打印配置）",
        status: "尚未实现（等待 Excel 修改能力验收）",
    },
    MenuEntry {
        number: "4",
        label: "按配置打印预校验（打印配置）",
        status: "Rust 入口",
    },
    MenuEntry {
        number: "5",
        label: "按配置打印 → 导出 PDF（打印配置）",
        status: "不迁移",
    },
];
const CONVERT: &[MenuEntry] = &[
    MenuEntry {
        number: "1",
        label: "批量重命名文件（重命名配置）",
        status: "Rust 入口",
    },
    MenuEntry {
        number: "2",
        label: "批量 Excel 格式转换（全局配置）",
        status: "尚未实现（尚未验收并接通）",
    },
    MenuEntry {
        number: "3",
        label: "批量 Word 格式转换（全局配置）",
        status: "尚未实现（尚未验收并接通）",
    },
    MenuEntry {
        number: "4",
        label: "批量修改 Sheet 名（重命名配置）",
        status: "尚未实现（等待 Excel 修改能力验收）",
    },
    MenuEntry {
        number: "5",
        label: "根据文件内容重命名",
        status: "Rust 入口",
    },
    MenuEntry {
        number: "6",
        label: "取消根据内容重命名前缀",
        status: "Rust 入口",
    },
    MenuEntry {
        number: "7",
        label: "批量局部映射重命名文件",
        status: "Rust 入口",
    },
    MenuEntry {
        number: "8",
        label: "提交缺失检查",
        status: "Rust 入口",
    },
    MenuEntry {
        number: "9",
        label: "文件排序前缀",
        status: "Rust 入口",
    },
    MenuEntry {
        number: "10",
        label: "取消文件排序前缀",
        status: "Rust 入口",
    },
];
const CONFIG: &[MenuEntry] = &[
    MenuEntry {
        number: "1",
        label: "初始化新配置（现有配置修复功能尚未实现）",
        status: "Rust 入口（仅新建，不覆盖已有文件）",
    },
    MenuEntry {
        number: "2",
        label: "查看当前配置（重新读取磁盘）",
        status: "Rust 入口",
    },
];
const SUMMARY: &[MenuEntry] = &[
    MenuEntry {
        number: "1",
        label: "按使用区域汇总（全量，自动路由）",
        status: "Rust 入口",
    },
    MenuEntry {
        number: "2",
        label: "按批注汇总（模板+源文件，自动路由）",
        status: "Rust 入口",
    },
];
const PENDING: &[MenuEntry] = &[
    MenuEntry {
        number: "1",
        label: "拆分村镇银行数据",
        status: "尚未实现（计划接口，未接通执行）",
    },
    MenuEntry {
        number: "2",
        label: "机构名称标准化（机构映射表）",
        status: "尚未实现（等待 Excel 修改能力验收）",
    },
    MenuEntry {
        number: "3",
        label: "删除分机构表的外资行（机构映射表）",
        status: "尚未实现（计划接口，未接通执行）",
    },
    MenuEntry {
        number: "4",
        label: "外汇页眉修改",
        status: "尚未实现（计划接口，未接通执行）",
    },
    MenuEntry {
        number: "5",
        label: "地区总分校验",
        status: "Rust 入口",
    },
    MenuEntry {
        number: "6",
        label: "提取工作表数据（工作表提取）",
        status: "尚未实现完整功能（仅有有限候选）",
    },
    MenuEntry {
        number: "7",
        label: "涉农贷款比上月修正",
        status: "尚未实现（计划接口，未接通执行）",
    },
    MenuEntry {
        number: "8",
        label: "季度涉农贷款比上月修正及分地区核对",
        status: "尚未实现（计划接口，未接通执行）",
    },
];
const ARCHIVE: &[MenuEntry] = &[
    MenuEntry {
        number: "1",
        label: "按日期归档（预览）",
        status: "Rust 入口",
    },
    MenuEntry {
        number: "2",
        label: "按类型归档（预览，归档类型配置）",
        status: "Rust 入口",
    },
    MenuEntry {
        number: "3",
        label: "执行归档计划（归档计划）",
        status: "Rust 入口",
    },
    MenuEntry {
        number: "4",
        label: "按日期归档（直接复制）",
        status: "Rust 入口",
    },
    MenuEntry {
        number: "5",
        label: "按类型归档（直接复制，归档类型配置）",
        status: "Rust 入口",
    },
];
const QUICK: &[MenuEntry] = &[
    MenuEntry {
        number: "1",
        label: "打开配置文件（config.xlsx）",
        status: "Rust 入口",
    },
    MenuEntry {
        number: "2",
        label: "打开 output 最新文件",
        status: "Rust 入口",
    },
    MenuEntry {
        number: "3",
        label: "打开 output 文件夹",
        status: "Rust 入口",
    },
];

pub fn run(mut dispatch: impl FnMut(Vec<String>) -> Result<()>) -> Result<()> {
    loop {
        println!("{MAIN_MENU}");
        let Some(group) = read_choice(
            "主菜单选择> ",
            &["0", "1", "2", "3", "4", "5", "6", "7", "8", "9"],
        )?
        else {
            println!("已关闭输入，退出菜单。");
            return Ok(());
        };
        if group == "0" {
            println!("再见。");
            return Ok(());
        }
        let (title, items) = match group.as_str() {
            "1" => ("时序工具", TIMELINE),
            "2" => ("去重工具", DEDUP),
            "3" => ("打印工具", PRINT),
            "4" => ("转换工具", CONVERT),
            "5" => ("配置工具", CONFIG),
            "6" => ("汇总工具", SUMMARY),
            "7" => ("一二批处理工具", PENDING),
            "8" => ("归档工具", ARCHIVE),
            "9" => ("快捷工具", QUICK),
            _ => unreachable!(),
        };
        let Some(selection) = read_submenu(title, items)? else {
            continue;
        };
        match build_command(&group, &selection) {
            Ok(Some(arguments)) => {
                println!("\n[Rust 入口] 正在执行已接通的命令；请查看命令输出和结果路径。\n");
                match dispatch(arguments) {
                    Ok(()) => println!("\n[Rust 命令已返回] 请按上方结果确认输出状态。"),
                    Err(error) => eprintln!("\n[执行失败] {error:#}"),
                }
            }
            Ok(None) => {
                let status = items
                    .iter()
                    .find(|item| item.number == selection)
                    .map(|item| item.status)
                    .unwrap_or("尚未实现");
                println!("\n[尚未接通] {status}；此菜单项未执行，也未报告成功。\n");
            }
            Err(error) => println!("\n[输入/运行错误] {error:#}\n"),
        }
    }
}

fn read_submenu<'a>(title: &str, items: &'a [MenuEntry]) -> Result<Option<String>> {
    loop {
        println!("\n------ {title} ------");
        for item in items {
            println!(" {:>2}  {}  [{}]", item.number, item.label, item.status);
        }
        println!("  0  返回");
        let mut allowed = items.iter().map(|item| item.number).collect::<Vec<_>>();
        allowed.push("0");
        let Some(choice) = read_choice("子菜单选择> ", &allowed)? else {
            return Ok(None);
        };
        if choice == "0" {
            return Ok(None);
        }
        return Ok(Some(choice));
    }
}

fn read_choice(prompt: &str, valid: &[&str]) -> Result<Option<String>> {
    loop {
        let Some(value) = read_line(prompt)? else {
            return Ok(None);
        };
        if valid.contains(&value.as_str()) {
            return Ok(Some(value));
        }
        println!("输入无效，请输入 {}。", valid.join("/"));
    }
}

fn read_line(prompt: &str) -> Result<Option<String>> {
    print!("{prompt}");
    io::stdout().flush()?;
    let mut value = String::new();
    if io::stdin().read_line(&mut value)? == 0 {
        return Ok(None);
    }
    Ok(Some(value.trim().to_owned()))
}

fn build_command(group: &str, selection: &str) -> Result<Option<Vec<String>>> {
    let command = match (group, selection) {
        ("1", "1") => {
            let Some(config) = prompt_config()? else {
                return Ok(None);
            };
            let Some(template) = prompt_existing_path("模板工作簿路径（输入 0 取消）> ", false)?
            else {
                return Ok(None);
            };
            let Some(sources) =
                prompt_paths("源工作簿路径（可用分号分隔多个，输入 0 取消）> ", false)?
            else {
                return Ok(None);
            };
            let mut args = vec!["timeline-compare".into()];
            push_option(&mut args, "--config", config);
            push_option(&mut args, "--template", template);
            push_paths(&mut args, sources);
            args
        }
        ("1", "3" | "4") => {
            let Some(config) = prompt_config()? else {
                return Ok(None);
            };
            let Some(sources) =
                prompt_paths("源工作簿路径（可用分号分隔多个，输入 0 取消）> ", false)?
            else {
                return Ok(None);
            };
            let command = if selection == "3" {
                "wide-summary-row-suffix"
            } else {
                "wide-summary"
            };
            let mut args = vec![command.into()];
            push_option(&mut args, "--config", config);
            push_paths(&mut args, sources);
            args
        }
        ("1", "5") => {
            let Some(config) = prompt_config()? else {
                return Ok(None);
            };
            let mut args = vec!["timeline-precheck".into()];
            push_option(&mut args, "--config", config);
            args
        }
        ("2", "6") => {
            let Some(config) = prompt_config()? else {
                return Ok(None);
            };
            let mut args = vec!["dedup-precheck".into()];
            push_option(&mut args, "--config", config);
            args
        }
        ("3", "4") => {
            let Some(config) = prompt_config()? else {
                return Ok(None);
            };
            let mut args = vec!["print-config-precheck".into()];
            push_option(&mut args, "--config", config);
            args
        }
        ("4", "1") => {
            let Some(config) = prompt_config()? else {
                return Ok(None);
            };
            let Some(paths) = prompt_paths(
                "待重命名文件或目录（可用分号分隔多个，输入 0 取消）> ",
                true,
            )?
            else {
                return Ok(None);
            };
            if !confirm("将按配置改名所选文件；是否继续？")? {
                return Ok(None);
            }
            let mut args = vec!["rename-batch".into()];
            push_option(&mut args, "--config", config);
            push_paths(&mut args, paths);
            args
        }
        ("4", "5" | "6" | "7" | "9" | "10") => {
            let Some(config) = prompt_config()? else {
                return Ok(None);
            };
            let Some(paths) =
                prompt_paths("待处理文件路径（可用分号分隔多个，输入 0 取消）> ", false)?
            else {
                return Ok(None);
            };
            if !confirm("此操作会更改所选文件的名称，但不写入工作簿；是否继续？")?
            {
                return Ok(None);
            }
            let command = match selection {
                "5" => "rename-by-content",
                "6" => "unrename-by-content",
                "7" => "rename-local-map",
                "9" => "file-sort-prefix",
                "10" => "remove-file-sort-prefix",
                _ => unreachable!(),
            };
            let mut args = vec![command.into()];
            push_option(&mut args, "--config", config);
            push_paths(&mut args, paths);
            args
        }
        ("4", "8") => {
            let Some(config) = prompt_config()? else {
                return Ok(None);
            };
            let Some(paths) =
                prompt_paths("待检查文件或目录（可用分号分隔多个，输入 0 取消）> ", true)?
            else {
                return Ok(None);
            };
            let mut args = vec!["submission-check".into()];
            push_option(&mut args, "--config", config);
            push_paths(&mut args, paths);
            args
        }
        ("5", "1") => {
            let path = prompt_new_config_path()?;
            let Some(path) = path else { return Ok(None) };
            let mut args = vec!["config-init".into()];
            push_option(&mut args, "--path", path);
            args
        }
        ("5", "2") => {
            let Some(config) = prompt_config()? else {
                return Ok(None);
            };
            let mut args = vec!["config-view".into()];
            push_option(&mut args, "--config", config);
            args
        }
        ("6", "1") => {
            let Some(sources) =
                prompt_paths("源工作簿路径（可用分号分隔多个，输入 0 取消）> ", false)?
            else {
                return Ok(None);
            };
            let Some(output) = prompt_output_dir()? else {
                return Ok(None);
            };
            let mut args = vec!["summary-used-range".into()];
            push_option(&mut args, "--output-dir", output);
            push_paths(&mut args, sources);
            args
        }
        ("6", "2") => {
            let Some(template) = prompt_existing_path("汇总模板路径（输入 0 取消）> ", false)?
            else {
                return Ok(None);
            };
            let Some(sources) =
                prompt_paths("源工作簿路径（可用分号分隔多个，输入 0 取消）> ", false)?
            else {
                return Ok(None);
            };
            let Some(output) = prompt_output_dir()? else {
                return Ok(None);
            };
            let mut args = vec!["summary-comments".into()];
            push_option(&mut args, "--template", template);
            push_option(&mut args, "--output-dir", output);
            push_paths(&mut args, sources);
            args
        }
        ("7", "5") => {
            let Some(paths) =
                prompt_paths("待检查工作簿路径（可用分号分隔多个，输入 0 取消）> ", false)?
            else {
                return Ok(None);
            };
            let mut args = vec!["region-sum-check".into()];
            push_paths(&mut args, paths);
            args
        }
        ("8", "1" | "4") => {
            let Some(config) = prompt_config()? else {
                return Ok(None);
            };
            let Some(target) = prompt_target_root()? else {
                return Ok(None);
            };
            let Some(sources) =
                prompt_paths("待归档文件或目录（可用分号分隔多个，输入 0 取消）> ", true)?
            else {
                return Ok(None);
            };
            let command = if selection == "1" {
                "archive-date-preview"
            } else {
                "archive-date-copy"
            };
            if selection == "4"
                && !confirm("将把所选文件复制到归档目录（不删除源文件）；是否继续？")?
            {
                return Ok(None);
            }
            let mut args = vec![command.into()];
            push_option(&mut args, "--config", config);
            push_option(&mut args, "--target-root", target);
            push_option(&mut args, "--output-dir", PathBuf::from("output"));
            push_paths(&mut args, sources);
            args
        }
        ("8", "2" | "5") => {
            let Some(config) = prompt_config()? else {
                return Ok(None);
            };
            let Some(target) = prompt_target_root()? else {
                return Ok(None);
            };
            let Some(sources) =
                prompt_paths("待归档文件或目录（可用分号分隔多个，输入 0 取消）> ", true)?
            else {
                return Ok(None);
            };
            let command = if selection == "2" {
                "archive-type-preview"
            } else {
                "archive-type-copy"
            };
            if selection == "5"
                && !confirm("将把所选文件复制到归档目录（不删除源文件）；是否继续？")?
            {
                return Ok(None);
            }
            let mut args = vec![command.into()];
            push_option(&mut args, "--config", config);
            push_option(&mut args, "--target-root", target);
            push_option(&mut args, "--output-dir", PathBuf::from("output"));
            push_paths(&mut args, sources);
            args
        }
        ("8", "3") => {
            let Some(plan) = prompt_existing_path("归档计划路径（输入 0 取消）> ", false)?
            else {
                return Ok(None);
            };
            if !confirm("执行归档计划会复制文件；不会删除源文件。是否继续？")?
            {
                return Ok(None);
            }
            let mut args = vec!["archive-execute-plan".into()];
            push_option(&mut args, "--plan", plan);
            push_option(&mut args, "--output-dir", PathBuf::from("output"));
            args
        }
        ("9", "1" | "2" | "3") => {
            let Some(config) = prompt_config()? else {
                return Ok(None);
            };
            let command = match selection {
                "1" => "open-config",
                "2" => "open-latest-output",
                "3" => "open-output-dir",
                _ => unreachable!(),
            };
            let mut args = vec![command.into()];
            push_option(&mut args, "--config", config);
            args
        }
        _ => return Ok(None),
    };
    Ok(Some(command))
}

fn default_config_path() -> PathBuf {
    if let Some(path) = std::env::var_os("PYTOOLS_CONFIG_PATH") {
        return PathBuf::from(path);
    }
    let candidates = [
        std::env::current_exe().ok().and_then(|path| {
            path.parent()
                .map(|parent| parent.join("config/config.xlsx"))
        }),
        std::env::current_dir()
            .ok()
            .map(|path| path.join("pytools-core/config/config.xlsx")),
        std::env::current_dir()
            .ok()
            .map(|path| path.join("config/config.xlsx")),
    ];
    candidates
        .into_iter()
        .flatten()
        .find(|path| path.is_file())
        .or_else(|| {
            std::env::current_exe().ok().and_then(|path| {
                path.parent()
                    .map(|parent| parent.join("config/config.xlsx"))
            })
        })
        .unwrap_or_else(|| PathBuf::from("config/config.xlsx"))
}

fn prompt_config() -> Result<Option<PathBuf>> {
    let default = default_config_path();
    let Some(value) = read_line(&format!(
        "Rust 配置文件路径 [默认 {}]（输入 0 取消）> ",
        default.display()
    ))?
    else {
        return Ok(None);
    };
    if value == "0" {
        return Ok(None);
    }
    let path = if value.is_empty() {
        default
    } else {
        PathBuf::from(unquote(&value))
    };
    if !path.is_file() {
        bail!(
            "配置文件不存在：{}；可先在 5-1 新建配置模板",
            path.display()
        );
    }
    Ok(Some(path))
}

fn prompt_new_config_path() -> Result<Option<PathBuf>> {
    let default = default_config_path();
    let Some(value) = read_line(&format!(
        "新配置文件路径 [默认 {}]（输入 0 取消）> ",
        default.display()
    ))?
    else {
        return Ok(None);
    };
    if value == "0" {
        return Ok(None);
    }
    let path = if value.is_empty() {
        default
    } else {
        PathBuf::from(unquote(&value))
    };
    if path.exists() {
        bail!("目标已存在，不会覆盖：{}", path.display());
    }
    Ok(Some(path))
}

fn prompt_existing_path(prompt: &str, allow_directory: bool) -> Result<Option<PathBuf>> {
    let Some(value) = read_line(prompt)? else {
        return Ok(None);
    };
    if value.is_empty() || value == "0" {
        return Ok(None);
    }
    let path = PathBuf::from(unquote(&value));
    if !path.is_file() && !(allow_directory && path.is_dir()) {
        bail!("路径不存在或类型不符：{}", path.display());
    }
    Ok(Some(path))
}

fn prompt_paths(prompt: &str, allow_directories: bool) -> Result<Option<Vec<PathBuf>>> {
    let Some(value) = read_line(prompt)? else {
        return Ok(None);
    };
    if value.is_empty() || value == "0" {
        return Ok(None);
    }
    let paths = value
        .replace('；', ";")
        .split(';')
        .map(|item| PathBuf::from(unquote(item.trim())))
        .filter(|path| !path.as_os_str().is_empty())
        .collect::<Vec<_>>();
    if paths.is_empty() {
        bail!("至少输入一个路径");
    }
    for path in &paths {
        if !path.is_file() && !(allow_directories && path.is_dir()) {
            bail!("路径不存在或类型不符：{}", path.display());
        }
    }
    Ok(Some(paths))
}

fn prompt_target_root() -> Result<Option<PathBuf>> {
    let Some(value) = read_line("归档目标目录（可输入新目录，输入 0 取消）> ")?
    else {
        return Ok(None);
    };
    if value.is_empty() || value == "0" {
        return Ok(None);
    }
    Ok(Some(PathBuf::from(unquote(&value))))
}

fn prompt_output_dir() -> Result<Option<PathBuf>> {
    let Some(value) = read_line("输出目录 [默认 ./output]（输入 0 取消）> ")? else {
        return Ok(None);
    };
    if value == "0" {
        return Ok(None);
    }
    Ok(Some(if value.is_empty() {
        PathBuf::from("output")
    } else {
        PathBuf::from(unquote(&value))
    }))
}

fn confirm(prompt: &str) -> Result<bool> {
    let Some(value) = read_line(&format!("{prompt} [y/N] "))? else {
        return Ok(false);
    };
    Ok(matches!(value.to_ascii_lowercase().as_str(), "y" | "yes"))
}

fn push_option(args: &mut Vec<String>, name: &str, value: PathBuf) {
    args.push(name.to_owned());
    args.push(value.to_string_lossy().into_owned());
}

fn push_paths(args: &mut Vec<String>, paths: Vec<PathBuf>) {
    args.extend(
        paths
            .into_iter()
            .map(|path| path.to_string_lossy().into_owned()),
    );
}

fn unquote(value: &str) -> String {
    value.trim().trim_matches('"').trim_matches('\'').to_owned()
}
