use pytools_core::archive_ops::{
    archive_by_date_copy, archive_by_date_preview, archive_by_type_copy, archive_by_type_preview,
    execute_archive_plan, ArchiveResult,
};
use serde_json::{json, Value};
use std::fs;
use std::path::{Path, PathBuf};
use std::process::Command;
use std::time::{SystemTime, UNIX_EPOCH};

fn cases() -> Value {
    serde_json::from_str(include_str!("fixtures/archive_cases.json")).unwrap()
}

fn baseline() -> Value {
    serde_json::from_str(include_str!(
        "../probe/reports/archive_python_baseline.json"
    ))
    .unwrap()
}

struct TempDir(PathBuf);

impl TempDir {
    fn new(label: &str) -> Self {
        let unique = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .unwrap()
            .as_nanos();
        let path = std::env::temp_dir().join(format!(
            "pytools-archive-{label}-{}-{unique}",
            std::process::id()
        ));
        fs::create_dir_all(&path).unwrap();
        Self(path)
    }
}

impl Drop for TempDir {
    fn drop(&mut self) {
        let _ = fs::remove_dir_all(&self.0);
    }
}

fn setup(root: &Path, cases: &Value, existing_target: Option<&str>) -> (PathBuf, Vec<PathBuf>) {
    let input_dir = root.join("input");
    fs::create_dir_all(&input_dir).unwrap();
    for name in cases["input_names"].as_array().unwrap() {
        let relative = name.as_str().unwrap();
        let path = input_dir.join(relative);
        fs::create_dir_all(path.parent().unwrap()).unwrap();
        let bytes = if path
            .file_name()
            .unwrap_or_default()
            .to_string_lossy()
            .starts_with("~$")
        {
            b"temporary".to_vec()
        } else {
            format!("fixture:{relative}").into_bytes()
        };
        fs::write(path, bytes).unwrap();
    }
    let missing = input_dir.join(cases["missing_name"].as_str().unwrap());
    if let Some(relative) = existing_target {
        let target = root.join("target").join(relative);
        fs::create_dir_all(target.parent().unwrap()).unwrap();
        fs::write(target, b"occupied target").unwrap();
    }
    (input_dir.clone(), vec![input_dir, missing])
}

fn relative(path: &str, root: &Path) -> String {
    if path.is_empty() {
        return String::new();
    }
    Path::new(path)
        .strip_prefix(root)
        .unwrap_or_else(|error| panic!("{} is outside {}: {error}", path, root.display()))
        .to_string_lossy()
        .replace('\\', "/")
}

fn preview_rows(result: &ArchiveResult, root: &Path) -> Vec<Value> {
    result
        .rows
        .iter()
        .map(|row| {
            json!({
                "source": relative(&row.source_path, root),
                "file_name": row.file_name,
                "mode": row.mode,
                "category": row.category,
                "target_dir": relative(&row.target_dir, root),
                "target": relative(&row.target_path, root),
                "status": row.status,
                "note": row.note,
            })
        })
        .collect()
}

fn copy_rows(result: &ArchiveResult, root: &Path) -> Vec<Value> {
    result
        .rows
        .iter()
        .map(|row| {
            json!({
                "source": relative(&row.source_path, root),
                "planned_target": relative(&row.planned_target_path, root),
                "actual_target": if row.status == "已复制" { relative(&row.target_path, root) } else { String::new() },
                "status": row.status,
                "note": row.note,
            })
        })
        .collect()
}

fn assert_preview(name: &str, result: &ArchiveResult, root: &Path, baseline: &Value) {
    let expected = &baseline[name];
    let stats = &expected["stats"];
    assert_eq!(
        result.files,
        stats["files"].as_u64().unwrap() as usize,
        "{name} files"
    );
    assert_eq!(
        result.invalid,
        stats["invalid"].as_u64().unwrap() as usize,
        "{name} invalid"
    );
    assert_eq!(
        result.rows.len(),
        stats["planned"].as_u64().unwrap() as usize,
        "{name} planned"
    );
    let skipped_extensions = result
        .rows
        .iter()
        .filter(|row| {
            matches!(
                row.status.as_str(),
                "跳过：后缀被排除" | "跳过：后缀不在归档范围"
            )
        })
        .count();
    assert_eq!(
        skipped_extensions,
        stats["skipped_ext"].as_u64().unwrap() as usize,
        "{name} extension filter"
    );
    assert_eq!(
        preview_rows(result, root),
        expected["rows"].as_array().unwrap().clone(),
        "{name} rows"
    );
    assert!(result.output.is_file(), "{name} report must be written");
}

fn assert_copy(name: &str, result: &ArchiveResult, root: &Path, baseline: &Value) {
    let expected = &baseline[name];
    let stats = &expected["stats"];
    assert_eq!(
        result.copied,
        stats["copied"].as_u64().unwrap() as usize,
        "{name} copied"
    );
    assert_eq!(
        result.skipped,
        stats["skip"].as_u64().unwrap() as usize,
        "{name} skipped"
    );
    assert_eq!(
        result.failed,
        stats["failed"].as_u64().unwrap() as usize,
        "{name} failed"
    );
    assert_eq!(
        result.rows.len(),
        stats["rows"].as_u64().unwrap() as usize,
        "{name} rows"
    );
    assert_eq!(
        copy_rows(result, root),
        expected["rows"].as_array().unwrap().clone(),
        "{name} report rows"
    );
    assert!(result.output.is_file(), "{name} report must be written");
    for row in result.rows.iter().filter(|row| row.status == "已复制") {
        let source = Path::new(&row.source_path);
        let target = Path::new(&row.target_path);
        assert_eq!(
            fs::read(source).unwrap(),
            fs::read(target).unwrap(),
            "{name} copied bytes"
        );
    }
}

#[test]
fn archive_previews_and_copy_workflows_match_python_baseline() {
    let cases = cases();
    let baseline = baseline();
    let config =
        Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/archive_type_config.xlsx");

    let root = TempDir::new("date-preview");
    let (_input, paths) = setup(&root.0, &cases, cases["existing_targets"]["date"].as_str());
    let result =
        archive_by_date_preview(&paths, &root.0.join("target"), &root.0.join("out"), None).unwrap();
    assert_preview("date_preview", &result, &root.0, &baseline);

    let root = TempDir::new("date-filtered-preview");
    let (_input, paths) = setup(&root.0, &cases, None);
    let result = archive_by_date_preview(
        &paths,
        &root.0.join("target"),
        &root.0.join("out"),
        Some(&config),
    )
    .unwrap();
    assert_preview("date_filtered_preview", &result, &root.0, &baseline);

    let root = TempDir::new("type-preview");
    let (_input, paths) = setup(&root.0, &cases, cases["existing_targets"]["type"].as_str());
    let result =
        archive_by_type_preview(&paths, &root.0.join("target"), &root.0.join("out"), &config)
            .unwrap();
    assert_preview("type_preview", &result, &root.0, &baseline);

    let root = TempDir::new("date-copy");
    let (input, _paths) = setup(&root.0, &cases, cases["existing_targets"]["date"].as_str());
    let result =
        archive_by_date_copy(&[input], &root.0.join("target"), &root.0.join("out"), None).unwrap();
    assert_copy("date_copy", &result, &root.0, &baseline);

    let root = TempDir::new("type-copy");
    let (input, _paths) = setup(&root.0, &cases, cases["existing_targets"]["type"].as_str());
    let result = archive_by_type_copy(
        &[input],
        &root.0.join("target"),
        &root.0.join("out"),
        &config,
    )
    .unwrap();
    assert_copy("type_copy", &result, &root.0, &baseline);

    let root = TempDir::new("plan-copy");
    let (_input, paths) = setup(&root.0, &cases, cases["existing_targets"]["date"].as_str());
    let plan = archive_by_date_preview(&paths, &root.0.join("target"), &root.0.join("plan"), None)
        .unwrap();
    let result = execute_archive_plan(&plan.output, &root.0.join("out")).unwrap();
    assert_copy("plan_copy", &result, &root.0, &baseline);
}

#[test]
fn date_archive_cli_accepts_the_python_compatible_extension_filter_config() {
    let cases = cases();
    let config =
        Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/archive_type_config.xlsx");
    let root = TempDir::new("date-preview-cli");
    let (_input, paths) = setup(&root.0, &cases, None);
    let output = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .arg("archive-date-preview")
        .arg("--target-root")
        .arg(root.0.join("target"))
        .arg("--config")
        .arg(config)
        .arg("--output-dir")
        .arg(root.0.join("out"))
        .args(paths)
        .output()
        .unwrap();
    assert!(
        output.status.success(),
        "CLI failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    let report: Value = serde_json::from_slice(&output.stdout).unwrap();
    assert_eq!(report["files"], 12);
    assert_eq!(report["invalid"], 1);
    assert_eq!(report["rows"].as_array().unwrap().len(), 13);
    assert_eq!(
        report["rows"]
            .as_array()
            .unwrap()
            .iter()
            .filter(|row| row["status"].as_str().unwrap().starts_with("跳过：后缀"))
            .count(),
        2
    );
    assert!(Path::new(report["output"].as_str().unwrap()).is_file());
}
