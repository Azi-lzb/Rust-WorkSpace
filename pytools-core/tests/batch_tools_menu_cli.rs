use anyhow::{Context, Result};
use excel_reader::read_workbook;
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
            "pytools-batch-menu-{}-{nonce}-{serial}",
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
fn interactive_batch_menu_runs_region_sum_check_through_real_cli() -> Result<()> {
    let root = TestDirectory::new()?;
    let config = create_config(root.path())?;
    let fixture_dir = Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/region_sum");
    let fixture_names = [
        "balanced.xlsx",
        "mismatch.xlsx",
        "no_target.xlsx",
        "non_numeric_missing.xlsx",
        "broken.xlsx",
    ];
    let inputs = fixture_names
        .iter()
        .map(|name| fixture_dir.join(name))
        .collect::<Vec<_>>();
    let original_inputs = inputs
        .iter()
        .map(|path| Ok((path.clone(), fs::read(path)?)))
        .collect::<Result<Vec<_>>>()?;
    let unsupported_xls = root.path().join("unsupported.xls");
    fs::write(&unsupported_xls, b"synthetic legacy workbook placeholder")?;
    let original_unsupported_xls = fs::read(&unsupported_xls)?;
    let input_list = inputs
        .iter()
        .map(|path| path.to_string_lossy())
        .collect::<Vec<_>>()
        .join(";")
        + ";"
        + &unsupported_xls.to_string_lossy();

    let output = run_menu(&config, &format!("7\n5\n1\n{input_list}\n0\n"))?;
    assert!(
        output.status.success(),
        "batch menu failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    let stdout = String::from_utf8_lossy(&output.stdout);
    assert!(stdout.contains("7  一二批处理工具"));
    assert!(stdout.contains("5  地区总分校验"));
    assert!(stdout.contains("已选择文件 6 个"));
    assert!(stdout.contains("先显式转换为 .xlsx"));
    assert!(stdout.contains(r#""files": 4"#));
    assert!(stdout.contains(r#""sheets": 3"#));
    assert!(stdout.contains(r#""errors": 1"#));
    assert!(stdout.contains(r#""failed_files": 1"#));

    let output_dir = root.path().join("app/output");
    let report_path = fs::read_dir(&output_dir)?
        .filter_map(|entry| entry.ok().map(|entry| entry.path()))
        .find(|path| {
            path.file_name()
                .and_then(|name| name.to_str())
                .is_some_and(|name| {
                    name.starts_with("地区数值核对结果_") && name.ends_with(".xlsx")
                })
        })
        .context("region sum result workbook was not created")?;
    let report = read_workbook(&report_path)?;
    let sheet = report
        .sheets
        .iter()
        .find(|sheet| sheet.name == "核对结果")
        .context("region sum result sheet missing")?;
    assert_eq!(sheet.values["B2"]["value"], "balanced.xlsx");
    assert_eq!(sheet.values["D3"]["value"], "错误");
    assert_eq!(sheet.values["B5"]["value"], "已检查文件数: 4");

    let log_dir = root.path().join("app/logs");
    assert!(fs::read_dir(log_dir)?.next().is_some());
    for (path, original) in original_inputs {
        assert_eq!(
            fs::read(&path)?,
            original,
            "input changed: {}",
            path.display()
        );
    }
    assert_eq!(fs::read(&unsupported_xls)?, original_unsupported_xls);
    Ok(())
}

fn create_config(root: &Path) -> Result<PathBuf> {
    let app = root.join("app");
    let config = app.join("config/config.xlsx");
    fs::create_dir_all(config.parent().context("config parent missing")?)?;
    fs::create_dir_all(app.join("input"))?;
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let global = workbook.add_worksheet();
    global.set_name("全局配置")?;
    global.write_string(0, 0, "配置项")?;
    global.write_string(0, 1, "值")?;
    let settings = [
        ("输入目录", app.join("input").to_string_lossy().into_owned()),
        (
            "输出目录",
            app.join("output").to_string_lossy().into_owned(),
        ),
        ("日志目录", app.join("logs").to_string_lossy().into_owned()),
    ];
    for (index, (key, value)) in settings.iter().enumerate() {
        global.write_string((index + 1) as u32, 0, *key)?;
        global.write_string((index + 1) as u32, 1, value)?;
    }
    workbook.save(&config)?;
    Ok(config)
}

fn run_menu(config: &Path, input: &str) -> Result<Output> {
    let mut child = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .arg("menu")
        .arg("--config")
        .arg(config)
        .stdin(Stdio::piped())
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .spawn()?;
    child
        .stdin
        .take()
        .context("menu stdin missing")?
        .write_all(input.as_bytes())?;
    Ok(child.wait_with_output()?)
}
