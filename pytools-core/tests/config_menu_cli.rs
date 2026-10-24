use anyhow::{Context, Result};
use excel_reader::read_workbook;
use rust_xlsxwriter::Workbook;
use serde_json::Value;
use std::{
    fs,
    io::Write,
    path::{Path, PathBuf},
    process::{Command, Output, Stdio},
    sync::atomic::{AtomicU64, Ordering},
    time::{SystemTime, UNIX_EPOCH},
};

static NEXT_TEST_DIRECTORY_ID: AtomicU64 = AtomicU64::new(0);

struct TestDirectory(PathBuf);

impl TestDirectory {
    fn new() -> Result<Self> {
        let nonce = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .context("system clock before Unix epoch")?
            .as_nanos();
        let serial = NEXT_TEST_DIRECTORY_ID.fetch_add(1, Ordering::Relaxed);
        let path = std::env::temp_dir().join(format!(
            "pytools-config-menu-{}-{nonce}-{serial}",
            std::process::id()
        ));
        fs::create_dir_all(&path)?;
        Ok(Self(path))
    }

    fn path(&self) -> &Path {
        &self.0
    }
}

impl Drop for TestDirectory {
    fn drop(&mut self) {
        let _ = fs::remove_dir_all(&self.0);
    }
}

#[test]
fn missing_configuration_bootstraps_through_menu_and_can_be_viewed() -> Result<()> {
    let root = TestDirectory::new()?;
    let config = root.path().join("app/config/config.xlsx");
    let input = "1\n5\n2\n1\n2\n0\n0\n";

    let output = run_menu(&config, input, "user")?;
    anyhow::ensure!(
        output.status.success(),
        "menu bootstrap failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    let stdout = String::from_utf8_lossy(&output.stdout);
    assert!(stdout.contains("5  配置工具"));
    assert!(stdout.contains("当前配置不可用，该工具暂不能运行"));
    assert!(stdout.contains("请先选择 5-1 初始化/修复配置表"));
    assert!(stdout.contains("已新建配置表"));
    assert!(stdout.contains("查看当前配置（重新读取磁盘）"));
    assert!(!stdout.contains("实际读取组件"));

    let workbook = read_workbook(&config)?;
    assert_eq!(workbook.sheets.len(), 11);
    assert_eq!(workbook.sheets[0].name, "全局配置");
    assert!(stdout.contains("源文件扩展名"));
    Ok(())
}

#[test]
fn existing_configuration_repair_menu_saves_a_plan_without_modifying_source() -> Result<()> {
    let root = TestDirectory::new()?;
    let config = root.path().join("app/config/config.xlsx");
    write_incomplete_config(&config)?;
    let source_before = fs::read(&config)?;

    let output = run_menu(&config, "5\n1\n0\n0\n", "developer")?;
    anyhow::ensure!(
        output.status.success(),
        "configuration repair menu failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    let stdout = String::from_utf8_lossy(&output.stdout);
    assert!(stdout.contains("[计划已生成]"));
    assert!(stdout.contains("源配置未修改，也未生成候选"));
    assert!(stdout.contains("实际读取组件: excel-reader / Calamine"));
    assert!(stdout.contains("修改引擎: 未执行"));

    let output_root = root.path().join("app/output");
    let action_dir = fs::read_dir(&output_root)?
        .map(|entry| entry.map(|entry| entry.path()))
        .collect::<std::io::Result<Vec<_>>>()?
        .into_iter()
        .find(|path| {
            path.file_name()
                .is_some_and(|name| name.to_string_lossy().starts_with("pytools-config-repair-"))
        })
        .context("configuration repair output directory missing")?;
    let plan_path = action_dir.join("repair-plan.json");
    let plan: Value = serde_json::from_slice(&fs::read(&plan_path)?)?;
    assert!(plan["operations"]
        .as_array()
        .is_some_and(|items| !items.is_empty()));
    assert!(!action_dir.join("config-candidate.xlsx").exists());
    assert_eq!(fs::read(&config)?, source_before);
    Ok(())
}

fn write_incomplete_config(path: &Path) -> Result<()> {
    fs::create_dir_all(path.parent().context("configuration parent missing")?)?;
    let mut workbook = Workbook::new();
    let global = workbook.add_worksheet();
    global.set_name("全局配置")?;
    global.write_string(0, 0, "旧键头")?;
    global.write_string(0, 1, "旧值头")?;
    for (row, (key, value)) in [
        ("输入目录", "input"),
        ("输出目录", "output"),
        ("日志目录", "logs"),
    ]
    .iter()
    .enumerate()
    {
        global.write_string((row + 1) as u32, 0, *key)?;
        global.write_string((row + 1) as u32, 1, *value)?;
    }
    workbook.save(path)?;
    Ok(())
}

fn run_menu(config: &Path, input: &str, log_view: &str) -> Result<Output> {
    let mut child = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .arg("menu")
        .arg("--config")
        .arg(config)
        .arg("--log-view")
        .arg(log_view)
        .stdin(Stdio::piped())
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .spawn()
        .context("start interactive configuration menu")?;
    child
        .stdin
        .take()
        .context("interactive stdin missing")?
        .write_all(input.as_bytes())
        .context("send interactive menu selections")?;
    child
        .wait_with_output()
        .context("wait for interactive menu")
}
