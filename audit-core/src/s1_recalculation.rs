//! Explicit S1 formula-cache recalculation through the existing calculation sidecar.
//!
//! The input to this module is already a separate formula audit copy. It never
//! edits the submitted source workbook or the manifest's original audit copy.

use anyhow::{Context, Result, bail, ensure};
use serde_json::Value;
use std::{
    fs,
    path::{Path, PathBuf},
    process::Command,
    time::{Instant, SystemTime, UNIX_EPOCH},
};

#[derive(Clone, Debug)]
pub struct S1RecalculationReport {
    pub engine: String,
    pub input: PathBuf,
    pub output: PathBuf,
    pub elapsed_ms: u128,
}

/// Require explicit Office/LibreOffice recalculation for a production S1 run.
/// Reading an existing formula cache remains available only as a clearly
/// identified diagnostic mode for Python/Rust parity and troubleshooting.
pub fn validate_s1_formula_cache_mode(
    sidecar_requested: bool,
    allow_saved_cache_diagnostics: bool,
) -> Result<()> {
    ensure!(
        sidecar_requested || allow_saved_cache_diagnostics,
        "S1-F01 正式审核必须通过 --recalculate-sidecar 强制 Office/LibreOffice 重算；仅作缓存诊断时显式使用 --allow-saved-formula-cache-diagnostics"
    );
    ensure!(
        !sidecar_requested || !allow_saved_cache_diagnostics,
        "--allow-saved-formula-cache-diagnostics 仅适用于不执行公式重算的诊断模式"
    );
    Ok(())
}

/// Recalculate an existing S1 formula copy through the separately built
/// `lo-convert` sidecar, then publish the validated output without overwriting.
///
/// `lo-convert` selects Microsoft Excel COM on Windows and LibreOffice Calc on
/// Linux. A Windows WPS fallback is rejected because S1's confirmed Windows
/// baseline is Microsoft Excel. The sidecar validates formula/cache presence;
/// the original copied workbook remains untouched.
pub fn recalculate_s1_formula_copy(
    sidecar_path: impl AsRef<Path>,
    input_path: impl AsRef<Path>,
    output_path: impl AsRef<Path>,
) -> Result<S1RecalculationReport> {
    let started = Instant::now();
    let sidecar_path = fs::canonicalize(sidecar_path.as_ref())
        .with_context(|| format!("读取公式重算 sidecar：{}", sidecar_path.as_ref().display()))?;
    ensure!(sidecar_path.is_file(), "公式重算 sidecar 不是文件");

    let input_path = fs::canonicalize(input_path.as_ref())
        .with_context(|| format!("读取 S1 公式审核副本：{}", input_path.as_ref().display()))?;
    ensure!(input_path.is_file(), "S1 公式审核副本不是文件");
    ensure!(
        is_xlsx(&input_path),
        "S1 真实公式重算目前仅支持 .xlsx 审核副本"
    );

    let requested_output = output_path.as_ref();
    ensure!(is_xlsx(requested_output), "S1 公式重算输出必须为 .xlsx");
    let output_parent = requested_output
        .parent()
        .filter(|parent| !parent.as_os_str().is_empty())
        .unwrap_or_else(|| Path::new("."));
    let output_parent = fs::canonicalize(output_parent)
        .with_context(|| format!("读取 S1 重算输出目录：{}", output_parent.display()))?;
    ensure!(output_parent.is_dir(), "S1 重算输出路径的父级不是目录");
    let output_name = requested_output
        .file_name()
        .context("S1 重算输出必须包含文件名")?;
    let output_path = output_parent.join(output_name);
    ensure!(
        !output_path.exists(),
        "拒绝覆盖 S1 重算输出：{}",
        output_path.display()
    );
    ensure!(
        !same_path(&input_path, &output_path),
        "S1 公式重算必须写入独立于输入副本的新文件"
    );

    let staging = StagingDirectory::create(&output_parent)?;
    let candidate = staging.path().join("recalculated.xlsx");
    let sidecar_input = sidecar_io_path(&input_path);
    let sidecar_output = sidecar_io_path(&candidate);
    let command = Command::new(&sidecar_path)
        .arg("recalculate")
        .arg("--input")
        .arg(&sidecar_input)
        .arg("--output")
        .arg(&sidecar_output)
        .arg("--engine")
        .arg("auto")
        .output()
        .with_context(|| format!("启动公式重算 sidecar：{}", sidecar_path.display()))?;
    if !command.status.success() {
        bail!(
            "S1 公式重算失败（sidecar exit={}）：{}{}",
            command.status,
            String::from_utf8_lossy(&command.stderr).trim(),
            stdout_suffix(&command.stdout)
        );
    }

    let report: Value =
        serde_json::from_slice(&command.stdout).context("解析公式重算 sidecar 的 JSON 结果")?;
    ensure!(
        report.get("status").and_then(Value::as_str) == Some("recalculated"),
        "公式重算 sidecar 未报告 recalculated 状态"
    );
    let engine = report
        .get("engine")
        .and_then(Value::as_str)
        .context("公式重算 sidecar 未报告实际引擎")?
        .to_owned();
    validate_s1_engine(&engine)?;
    ensure!(
        candidate.is_file() && fs::metadata(&candidate)?.len() > 0,
        "公式重算 sidecar 未生成有效输出工作簿"
    );

    // The staging directory is inside the destination directory, so a hard
    // link publishes atomically and fails if a concurrent run created the
    // destination. We do not fall back to an overwrite-capable rename/copy.
    fs::hard_link(&candidate, &output_path).with_context(|| {
        format!(
            "安全发布 S1 重算副本失败（输出不得覆盖）：{}",
            output_path.display()
        )
    })?;

    Ok(S1RecalculationReport {
        engine,
        input: input_path,
        output: output_path,
        elapsed_ms: started.elapsed().as_millis(),
    })
}

pub fn s1_recalculated_filename(index: usize, source_path: impl AsRef<Path>) -> Result<String> {
    ensure!(index > 0, "S1 重算副本序号必须从 1 开始");
    let name = source_path
        .as_ref()
        .file_stem()
        .and_then(|stem| stem.to_str())
        .filter(|stem| !stem.trim().is_empty())
        .unwrap_or("audit-copy");
    let safe = name
        .chars()
        .take(80)
        .map(|character| {
            if character.is_alphanumeric() || matches!(character, '-' | '_' | '.') {
                character
            } else {
                '_'
            }
        })
        .collect::<String>();
    Ok(format!("{index:04}-{safe}-recalculated.xlsx"))
}

fn validate_s1_engine(engine: &str) -> Result<()> {
    #[cfg(windows)]
    ensure!(
        engine == "Excel.Application",
        "S1 Windows 公式重算要求 Microsoft Excel COM；sidecar 实际使用“{engine}”，拒绝将其结果作为审核输入"
    );
    #[cfg(not(windows))]
    ensure!(
        engine == "LibreOffice",
        "S1 Linux 公式重算要求 LibreOffice Calc；sidecar 实际使用“{engine}”，拒绝将其结果作为审核输入"
    );
    Ok(())
}

fn is_xlsx(path: &Path) -> bool {
    path.extension()
        .and_then(|extension| extension.to_str())
        .is_some_and(|extension| extension.eq_ignore_ascii_case("xlsx"))
}

fn same_path(left: &Path, right: &Path) -> bool {
    #[cfg(windows)]
    {
        left.to_string_lossy()
            .eq_ignore_ascii_case(&right.to_string_lossy())
    }
    #[cfg(not(windows))]
    {
        left == right
    }
}

fn sidecar_io_path(path: &Path) -> PathBuf {
    #[cfg(windows)]
    {
        let path = path.to_string_lossy();
        if let Some(path) = path.strip_prefix(r"\\?\UNC\") {
            return PathBuf::from(format!(r"\\{path}"));
        }
        if let Some(path) = path.strip_prefix(r"\\?\") {
            return PathBuf::from(path);
        }
        PathBuf::from(path.as_ref())
    }
    #[cfg(not(windows))]
    {
        path.to_path_buf()
    }
}

fn stdout_suffix(stdout: &[u8]) -> String {
    let text = String::from_utf8_lossy(stdout).trim().to_owned();
    if text.is_empty() {
        String::new()
    } else {
        format!("；stdout：{text}")
    }
}

struct StagingDirectory(PathBuf);

impl StagingDirectory {
    fn create(parent: &Path) -> Result<Self> {
        let timestamp = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .context("读取系统时间")?
            .as_nanos();
        for attempt in 0..100 {
            let path = parent.join(format!(
                ".s1-recalculate-{}-{timestamp}-{attempt}",
                std::process::id()
            ));
            match fs::create_dir(&path) {
                Ok(()) => return Ok(Self(path)),
                Err(error) if error.kind() == std::io::ErrorKind::AlreadyExists => continue,
                Err(error) => {
                    return Err(error)
                        .with_context(|| format!("创建 S1 公式重算临时目录：{}", path.display()));
                }
            }
        }
        bail!("无法为 S1 公式重算创建唯一临时目录")
    }

    fn path(&self) -> &Path {
        &self.0
    }
}

impl Drop for StagingDirectory {
    fn drop(&mut self) {
        let _ = fs::remove_dir_all(&self.0);
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use excel_reader::CellValue;
    use excel_writer::Workbook;
    use std::time::UNIX_EPOCH;

    #[test]
    fn recalculated_filename_is_safe_and_keeps_original_stem() {
        assert_eq!(
            s1_recalculated_filename(2, "个人贷款:2026.xlsx").unwrap(),
            "0002-个人贷款_2026-recalculated.xlsx"
        );
        assert!(s1_recalculated_filename(0, "sample.xlsx").is_err());
    }

    #[test]
    fn engine_must_match_the_confirmed_platform_baseline() {
        #[cfg(windows)]
        {
            assert!(validate_s1_engine("Excel.Application").is_ok());
            assert!(validate_s1_engine("KET.Application").is_err());
        }
        #[cfg(not(windows))]
        {
            assert!(validate_s1_engine("LibreOffice").is_ok());
            assert!(validate_s1_engine("Excel.Application").is_err());
        }
    }

    #[test]
    fn cached_formula_values_require_an_explicit_diagnostic_mode() {
        let error = validate_s1_formula_cache_mode(false, false)
            .expect_err("stored caches must not run as a production S1 audit");
        assert!(error.to_string().contains("强制 Office/LibreOffice 重算"));
        assert!(validate_s1_formula_cache_mode(false, true).is_ok());
        assert!(validate_s1_formula_cache_mode(true, false).is_ok());
        assert!(validate_s1_formula_cache_mode(true, true).is_err());
    }

    #[cfg(windows)]
    #[test]
    fn sidecar_paths_remove_windows_verbatim_prefix_for_com() {
        assert_eq!(
            sidecar_io_path(Path::new(r"\\?\C:\temp\audit.xlsx")),
            PathBuf::from(r"C:\temp\audit.xlsx")
        );
        assert_eq!(
            sidecar_io_path(Path::new(r"\\?\UNC\server\share\audit.xlsx")),
            PathBuf::from(r"\\server\share\audit.xlsx")
        );
    }

    #[test]
    fn refuses_to_overwrite_an_existing_output() -> Result<()> {
        let root = TestDirectory::create()?;
        let input = root.join("audit-copy.xlsx");
        let output = root.join("recalculated.xlsx");
        let sidecar = root.join("sidecar.exe");
        fs::write(&input, b"source-copy")?;
        fs::write(&sidecar, b"test-executable")?;
        fs::write(&output, b"existing-output")?;
        let input_before = fs::read(&input)?;

        let error = recalculate_s1_formula_copy(&sidecar, &input, &output)
            .expect_err("existing output must be refused");
        assert!(format!("{error:#}").contains("拒绝覆盖 S1 重算输出"));
        assert_eq!(fs::read(&input)?, input_before);
        assert_eq!(fs::read(&output)?, b"existing-output");
        Ok(())
    }

    #[test]
    #[ignore = "requires AUDIT_S1_RECALC_TEST_SIDECAR and a real supported calculation engine"]
    fn explicit_sidecar_recalculates_synthetic_formula_copy_without_changing_input() -> Result<()> {
        let sidecar = std::env::var_os("AUDIT_S1_RECALC_TEST_SIDECAR")
            .map(PathBuf::from)
            .context(
                "AUDIT_S1_RECALC_TEST_SIDECAR is required for the explicit integration test",
            )?;
        let root = TestDirectory::create()?;
        let input = root.join("synthetic-formula-copy.xlsx");
        let output = root.join("synthetic-formula-recalculated.xlsx");
        let mut workbook = Workbook::new();
        let sheet = workbook.add_worksheet();
        sheet.write_number(0, 0, 2.0)?;
        sheet.write_number(0, 1, 3.0)?;
        sheet.write_formula(0, 2, "=SUM(A1:B1)")?;
        workbook.save(&input)?;
        let input_before = fs::read(&input)?;

        let recalculation = recalculate_s1_formula_copy(&sidecar, &input, &output);
        assert_eq!(fs::read(&input)?, input_before);
        if recalculation.is_err() {
            assert!(
                !output.exists(),
                "failed calculation must not publish an output"
            );
        }
        let report = recalculation?;
        assert_eq!(report.output, fs::canonicalize(&output)?);
        assert!(output.is_file());
        assert_eq!(fs::read(&input)?, input_before);
        let values = excel_reader::read_value_columns(&output, 3)?;
        let calculated = values[0]
            .values
            .get("C1")
            .context("recalculated formula cache missing from C1")?;
        match calculated {
            CellValue::Number(value) | CellValue::DateTimeSerial(value) => {
                assert!((*value - 5.0).abs() < f64::EPSILON)
            }
            CellValue::Integer(value) => assert_eq!(*value, 5),
            other => panic!("unexpected recalculated C1 value: {other:?}"),
        }
        Ok(())
    }

    struct TestDirectory(PathBuf);

    impl TestDirectory {
        fn create() -> Result<Self> {
            let nonce = SystemTime::now().duration_since(UNIX_EPOCH)?.as_nanos();
            let root = std::env::temp_dir().join(format!(
                "audit_s1_recalculation_{}_{}",
                std::process::id(),
                nonce
            ));
            fs::create_dir(&root)?;
            Ok(Self(root))
        }
    }

    impl std::ops::Deref for TestDirectory {
        type Target = Path;

        fn deref(&self) -> &Self::Target {
            &self.0
        }
    }

    impl Drop for TestDirectory {
        fn drop(&mut self) {
            let _ = fs::remove_dir_all(&self.0);
        }
    }
}
