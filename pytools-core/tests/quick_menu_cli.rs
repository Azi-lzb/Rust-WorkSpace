use anyhow::{Context, Result};
use std::{
    fs,
    io::Write,
    path::{Path, PathBuf},
    process::{Command, Stdio},
    sync::atomic::{AtomicU64, Ordering},
    time::{SystemTime, UNIX_EPOCH},
};

static NEXT_TEMP_DIRECTORY: AtomicU64 = AtomicU64::new(0);

struct TestDirectory(PathBuf);

impl TestDirectory {
    fn new() -> Result<Self> {
        let nonce = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .context("system clock before Unix epoch")?
            .as_nanos();
        let serial = NEXT_TEMP_DIRECTORY.fetch_add(1, Ordering::Relaxed);
        let path = std::env::temp_dir().join(format!(
            "pytools-quick-menu-{}-{nonce}-{serial}",
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
fn main_menu_routes_all_three_quick_actions_without_starting_a_desktop_app() -> Result<()> {
    let root = TestDirectory::new()?;
    let blocked_parent = root.path().join("blocked-parent");
    fs::write(&blocked_parent, b"not a directory")?;
    let config = blocked_parent.join("config").join("config.xlsx");

    let mut child = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .args([
            "menu",
            "--config",
            config.to_string_lossy().as_ref(),
            "--log-view",
            "off",
        ])
        .stdin(Stdio::piped())
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .spawn()?;
    child
        .stdin
        .take()
        .context("menu stdin was not piped")?
        .write_all(b"9\n1\n9\n2\n9\n3\n0\n")?;
    let result = child.wait_with_output()?;
    assert!(
        result.status.success(),
        "quick menu exited unsuccessfully: {}",
        String::from_utf8_lossy(&result.stderr)
    );

    let stdout = String::from_utf8_lossy(&result.stdout);
    assert!(
        stdout.contains("  9  快捷工具"),
        "main menu missing option 9: {stdout}"
    );
    assert!(
        stdout.contains("------ 快捷工具 ------"),
        "quick menu missing: {stdout}"
    );
    assert!(
        stdout.contains("[失败] 打开配置文件失败:"),
        "config route missing: {stdout}"
    );
    assert!(
        stdout.contains("[提示] 未找到输出文件:"),
        "latest-output route missing: {stdout}"
    );
    assert!(
        stdout.contains("[失败] 打开输出文件夹失败:"),
        "output-folder route missing: {stdout}"
    );
    assert!(
        stdout.contains("再见。"),
        "quick action did not return to the main menu: {stdout}"
    );
    Ok(())
}

#[test]
fn launching_without_arguments_opens_the_legacy_numbered_main_menu() -> Result<()> {
    let mut child = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .stdin(Stdio::piped())
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .spawn()?;
    child
        .stdin
        .take()
        .context("menu stdin was not piped")?
        .write_all(b"0\n")?;
    let result = child.wait_with_output()?;
    assert!(
        result.status.success(),
        "direct menu launch failed: {}",
        String::from_utf8_lossy(&result.stderr)
    );

    let stdout = String::from_utf8_lossy(&result.stdout);
    for legacy_entry in [
        "======== PyTools 菜单 ========",
        "  1  时序工具",
        "  2  去重工具",
        "  3  打印工具",
        "  4  转换",
        "  5  配置工具",
        "  6  汇总工具",
        "  7  一二批处理工具",
        "  8  归档工具",
        "  9  快捷工具",
        "  0  退出",
        "再见。",
    ] {
        assert!(
            stdout.contains(legacy_entry),
            "direct launch is missing {legacy_entry:?}: {stdout}"
        );
    }
    Ok(())
}

#[test]
fn launching_without_arguments_honors_the_legacy_config_path_override() -> Result<()> {
    let root = TestDirectory::new()?;
    let config = root.path().join("selected config.xlsx");
    let mut child = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .env("PYTOOLS_CONFIG_PATH", &config)
        .stdin(Stdio::piped())
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .spawn()?;
    child
        .stdin
        .take()
        .context("menu stdin was not piped")?
        .write_all(b"0\n")?;
    let result = child.wait_with_output()?;
    assert!(
        result.status.success(),
        "direct menu launch failed: {}",
        String::from_utf8_lossy(&result.stderr)
    );

    let stdout = String::from_utf8_lossy(&result.stdout);
    assert!(
        stdout.contains(&config.to_string_lossy().to_string()),
        "menu did not use PYTOOLS_CONFIG_PATH: {stdout}"
    );
    assert!(stdout.contains("再见。"));
    Ok(())
}

#[test]
fn cli_help_lists_quick_tools_menu_category() -> Result<()> {
    let result = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .arg("--help")
        .output()?;
    assert!(
        result.status.success(),
        "pytools --help failed: {}",
        String::from_utf8_lossy(&result.stderr)
    );

    let usage = String::from_utf8_lossy(&result.stderr);
    assert!(
        usage.contains("9 (quick tools 1-3)"),
        "usage should list the quick tools menu category: {usage}"
    );
    Ok(())
}
