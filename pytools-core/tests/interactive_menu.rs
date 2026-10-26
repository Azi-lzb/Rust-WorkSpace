use pytools_core::config_admin::initialize_new_config;
use std::{
    fs,
    io::Write,
    process::{Command, Stdio},
    time::{SystemTime, UNIX_EPOCH},
};

fn temporary_directory(label: &str) -> std::path::PathBuf {
    let nonce = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .expect("system clock is after Unix epoch")
        .as_nanos();
    std::env::temp_dir().join(format!(
        "pytools-menu-{label}-{}-{nonce}",
        std::process::id()
    ))
}

#[test]
fn no_arguments_enters_the_legacy_numbered_menu_and_exits_cleanly() {
    let mut child = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .stdin(Stdio::piped())
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .spawn()
        .expect("start pytools menu");
    child
        .stdin
        .take()
        .expect("menu stdin")
        .write_all(b"0\n")
        .expect("send exit choice");
    let output = child.wait_with_output().expect("collect menu output");
    let stdout = String::from_utf8_lossy(&output.stdout);

    assert!(
        output.status.success(),
        "{}",
        String::from_utf8_lossy(&output.stderr)
    );
    assert!(stdout.contains("======== pytools 菜单 ========"));
    assert!(stdout.contains("  1  时序工具"));
    assert!(stdout.contains("  9  快捷工具"));
    assert!(stdout.contains("主菜单选择> 再见。"));
}

#[test]
fn config_view_is_reached_from_the_original_numeric_submenu() {
    let directory = temporary_directory("config-view");
    fs::create_dir_all(&directory).expect("create isolated test directory");
    let config = directory.join("config.xlsx");
    initialize_new_config(&config).expect("create a new config fixture");
    let before = fs::read(&config).expect("read fixture bytes before viewing");

    let mut child = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .stdin(Stdio::piped())
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .spawn()
        .expect("start pytools menu");
    let input = format!("5\n2\n{}\n0\n", config.display());
    child
        .stdin
        .take()
        .expect("menu stdin")
        .write_all(input.as_bytes())
        .expect("send config menu choices");
    let output = child.wait_with_output().expect("collect menu output");
    let stdout = String::from_utf8_lossy(&output.stdout);

    assert!(
        output.status.success(),
        "{}",
        String::from_utf8_lossy(&output.stderr)
    );
    assert!(stdout.contains("查看当前配置（重新读取磁盘）"));
    assert!(stdout.contains("全局配置"));
    assert!(stdout.contains("[Rust 入口] 正在执行"));
    assert!(stdout.contains("再见。"));
    assert!(
        config.is_file(),
        "the selected config must not be overwritten or removed"
    );
    assert_eq!(
        fs::read(&config).expect("read fixture bytes after viewing"),
        before
    );

    fs::remove_dir_all(directory).expect("remove only this test's temporary directory");
}

#[test]
fn config_initialization_uses_the_safe_new_file_command() {
    let directory = temporary_directory("config-init");
    fs::create_dir_all(&directory).expect("create isolated test directory");
    let config = directory.join("new-config.xlsx");

    let mut child = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .stdin(Stdio::piped())
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .spawn()
        .expect("start pytools menu");
    let input = format!("5\n1\n{}\n0\n", config.display());
    child
        .stdin
        .take()
        .expect("menu stdin")
        .write_all(input.as_bytes())
        .expect("send config initialization choices");
    let output = child.wait_with_output().expect("collect menu output");
    let stdout = String::from_utf8_lossy(&output.stdout);

    assert!(
        output.status.success(),
        "{}",
        String::from_utf8_lossy(&output.stderr)
    );
    assert!(stdout.contains("new configuration template created"));
    assert!(config.is_file());

    fs::remove_dir_all(directory).expect("remove only this test's temporary directory");
}

#[test]
fn used_range_summary_menu_route_creates_output_and_preserves_source() {
    let source = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("tests/fixtures/config_task_loaders.xlsx");
    let source_before = fs::read(&source).expect("read synthetic source bytes");
    let directory = temporary_directory("summary");
    let output_dir = directory.join("output");
    fs::create_dir_all(&directory).expect("create isolated test directory");

    let mut child = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .stdin(Stdio::piped())
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .spawn()
        .expect("start pytools menu");
    let input = format!("6\n1\n{}\n{}\n0\n", source.display(), output_dir.display());
    child
        .stdin
        .take()
        .expect("menu stdin")
        .write_all(input.as_bytes())
        .expect("send summary choices");
    let output = child.wait_with_output().expect("collect menu output");
    let stdout = String::from_utf8_lossy(&output.stdout);

    assert!(
        output.status.success(),
        "{}",
        String::from_utf8_lossy(&output.stderr)
    );
    assert!(stdout.contains("\"saved\": true"), "{stdout}");
    assert!(fs::read_dir(&output_dir)
        .expect("summary output directory exists")
        .any(|entry| entry.is_ok_and(|entry| entry
            .path()
            .extension()
            .is_some_and(|extension| extension == "xlsx"))));
    assert_eq!(
        fs::read(&source).expect("read synthetic source bytes after run"),
        source_before
    );

    fs::remove_dir_all(directory).expect("remove only this test's temporary directory");
}

#[test]
fn unavailable_menu_item_reports_that_it_did_not_run() {
    let mut child = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .stdin(Stdio::piped())
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .spawn()
        .expect("start pytools menu");
    child
        .stdin
        .take()
        .expect("menu stdin")
        .write_all("4\n2\n0\n".as_bytes())
        .expect("send unavailable conversion choice");
    let output = child.wait_with_output().expect("collect menu output");
    let stdout = String::from_utf8_lossy(&output.stdout);

    assert!(output.status.success());
    assert!(stdout.contains("尚未实现"));
    assert!(stdout.contains("[尚未接通]"));
    assert!(!stdout.contains("[Rust 命令已返回]"));
}
