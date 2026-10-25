use audit_core::s2_service::{run_s2_audit, S2ProgressEvent, S2RunRequest};
use excel_reader::read_workbook;
use serde_json::Value;
use std::{
    env, fs,
    path::{Path, PathBuf},
    process::Command,
    time::{SystemTime, UNIX_EPOCH},
};

#[test]
fn s2_service_runs_both_file_based_paths_and_keeps_inputs_unchanged() {
    let audit_root = Path::new(env!("CARGO_MANIFEST_DIR"));
    let fixture_root = audit_root.join("tests/fixtures");
    let configured_probe = env::var_os("AUDIT_S2_SERVICE_PROBE_DIR").map(PathBuf::from);
    let probe_root = configured_probe
        .clone()
        .unwrap_or_else(|| unique_test_dir());
    let should_cleanup = configured_probe.is_none();

    if probe_root.exists() && should_cleanup {
        fs::remove_dir_all(&probe_root).unwrap();
    }

    let input_root = probe_root.join("input");
    let current_dir = input_root.join("current");
    let previous_dir = input_root.join("previous");
    fs::create_dir_all(&current_dir).unwrap();
    fs::create_dir_all(&previous_dir).unwrap();

    let source_report = fixture_root.join("s2_import/9144#2026-06-30#x#20202#甲银行.xlsx");
    fs::copy(
        &source_report,
        current_dir.join("9144#2026-06-30#x#20202#甲银行.xlsx"),
    )
    .unwrap();
    fs::copy(
        &source_report,
        previous_dir.join("9144#2026-03-31#x#20202#甲银行.xlsx"),
    )
    .unwrap();

    let config_path = input_root.join("check_valid.xlsx");
    fs::copy(
        fixture_root.join("s2_config/check_valid.xlsx"),
        &config_path,
    )
    .unwrap();
    let invalid_config_path = input_root.join("check_gap_duplicate.xlsx");
    fs::copy(
        fixture_root.join("s2_config/check_gap_duplicate.xlsx"),
        &invalid_config_path,
    )
    .unwrap();
    let external_path = input_root.join("plain_layout.xlsx");
    fs::copy(
        fixture_root.join("s2_external_import/plain_layout.xlsx"),
        &external_path,
    )
    .unwrap();

    let before = input_snapshot(&input_root);
    for (scenario, external_workbook) in [
        ("without_external", None),
        ("with_external", Some(external_path.clone())),
    ] {
        let scenario_root = probe_root.join("rust").join(scenario);
        let output_directory = scenario_root.join("output");
        let request = S2RunRequest {
            current_directory: current_dir.clone(),
            previous_directory: previous_dir.clone(),
            config_path: config_path.clone(),
            output_directory,
            external_workbook,
        };
        let mut events: Vec<S2ProgressEvent> = Vec::new();
        let result = run_s2_audit(&request, Some(&mut |event| events.push(event))).unwrap();

        assert!(!result.analysis.current_records.is_empty());
        assert_eq!(
            result.analysis.summary.total_findings,
            result.analysis.findings.len()
        );
        assert!(result.report_path.is_file());
        let workbook = read_workbook(&result.report_path).unwrap();
        assert_eq!(
            workbook
                .sheets
                .iter()
                .map(|sheet| sheet.name.as_str())
                .collect::<Vec<_>>(),
            ["环比规则", "外部核对规则", "校验规则"]
        );

        assert_eq!(events.first().unwrap().stage, "config");
        assert_eq!(events.last().unwrap().stage, "complete");
        assert_eq!(events.last().unwrap().percent, 100.0);
        assert!(events
            .windows(2)
            .all(|pair| pair[0].percent <= pair[1].percent));
        assert_eq!(
            events.iter().any(|event| event.stage == "external_read"),
            scenario == "with_external"
        );

        fs::create_dir_all(&scenario_root).unwrap();
        fs::write(
            scenario_root.join("analysis.json"),
            serde_json::to_vec_pretty(&result.analysis).unwrap(),
        )
        .unwrap();
        fs::write(
            scenario_root.join("progress.json"),
            serde_json::to_vec_pretty(&events).unwrap(),
        )
        .unwrap();
        fs::write(
            scenario_root.join("report_path.txt"),
            result.report_path.to_string_lossy().as_bytes(),
        )
        .unwrap();
    }

    let scenario_root = probe_root.join("rust/invalid_config");
    let output_directory = scenario_root.join("output");
    let request = S2RunRequest {
        current_directory: current_dir,
        previous_directory: previous_dir,
        config_path: invalid_config_path,
        output_directory: output_directory.clone(),
        external_workbook: None,
    };
    let mut events = Vec::new();
    let error = run_s2_audit(&request, Some(&mut |event| events.push(event))).unwrap_err();
    assert_eq!(events.len(), 1);
    assert_eq!(events[0].stage, "config");
    assert!(!output_directory.exists());
    fs::create_dir_all(&scenario_root).unwrap();
    fs::write(scenario_root.join("error.txt"), error.to_string()).unwrap();
    fs::write(
        scenario_root.join("progress.json"),
        serde_json::to_vec_pretty(&events).unwrap(),
    )
    .unwrap();

    let cli_root = probe_root.join("rust_cli");
    for (scenario, config_name, external_workbook) in [
        ("without_external", "check_valid.xlsx", None),
        (
            "with_external",
            "check_valid.xlsx",
            Some(external_path.clone()),
        ),
        ("invalid_config", "check_gap_duplicate.xlsx", None),
    ] {
        let scenario_root = cli_root.join(scenario);
        let output_directory = scenario_root.join("output");
        let mut command = Command::new(env!("CARGO_BIN_EXE_open-data-audit"));
        command
            .arg("s2")
            .arg("--current-dir")
            .arg(input_root.join("current"))
            .arg("--previous-dir")
            .arg(input_root.join("previous"))
            .arg("--config")
            .arg(input_root.join(config_name))
            .arg("--output-dir")
            .arg(&output_directory);
        if let Some(path) = external_workbook {
            command.arg("--external").arg(path);
        }
        let output = command.output().unwrap();
        let stderr = String::from_utf8_lossy(&output.stderr);
        let mut cli_events = Vec::<S2ProgressEvent>::new();
        let mut cli_log_events = Vec::<Value>::new();
        let mut cli_errors = Vec::new();
        for line in stderr.lines() {
            if let Ok(value) = serde_json::from_str::<Value>(line) {
                if value["type"] == "progress" {
                    cli_log_events.push(value.clone());
                    cli_events.push(serde_json::from_value(value["event"].clone()).unwrap());
                }
            } else if let Some(error) = line.strip_prefix("open-data-audit: ") {
                cli_errors.push(error.to_owned());
            }
        }
        assert!(!cli_log_events.is_empty());
        assert!(cli_log_events.iter().all(|entry| {
            entry["engine"].is_string()
                && entry["elapsed_ms"].as_u64().is_some()
                && entry["stage_elapsed_ms"].as_u64().is_some()
        }));
        assert!(cli_log_events
            .windows(2)
            .all(|pair| { pair[0]["elapsed_ms"].as_u64() <= pair[1]["elapsed_ms"].as_u64() }));
        if scenario == "invalid_config" {
            assert!(cli_log_events
                .iter()
                .any(|entry| entry["event"]["stage"] == "config" && entry["engine"] == "Calamine"));
        } else {
            assert!(
                cli_log_events
                    .iter()
                    .any(|entry| entry["event"]["stage"] == "current"
                        && entry["engine"] == "Calamine")
            );
            assert!(cli_log_events
                .iter()
                .any(|entry| entry["event"]["stage"] == "rules"
                    && entry["engine"] == "Rust audit-core"));
            assert!(cli_log_events
                .iter()
                .any(|entry| entry["event"]["stage"] == "export"
                    && entry["engine"] == "rust_xlsxwriter"));
        }
        fs::create_dir_all(&scenario_root).unwrap();
        let cli_result = if scenario == "invalid_config" {
            assert_eq!(output.status.code(), Some(2));
            assert!(output.stdout.is_empty());
            assert!(!output_directory.exists());
            assert_eq!(cli_events.len(), 1);
            assert_eq!(cli_events[0].stage, "config");
            assert_eq!(cli_errors.len(), 1);
            serde_json::json!({"error": cli_errors[0]})
        } else {
            assert!(output.status.success(), "Audit CLI failed: {}", stderr);
            assert!(cli_errors.is_empty());
            let response: Value = serde_json::from_slice(&output.stdout).unwrap();
            assert_eq!(response["status"], "completed");
            let report_path = PathBuf::from(response["report_path"].as_str().unwrap());
            assert!(report_path.is_file());
            serde_json::json!({
                "status": response["status"],
                "report_path": report_path,
                "summary": response["summary"],
            })
        };
        fs::write(
            scenario_root.join("cli_result.json"),
            serde_json::to_vec_pretty(&cli_result).unwrap(),
        )
        .unwrap();
        fs::write(
            scenario_root.join("progress.json"),
            serde_json::to_vec_pretty(&cli_events).unwrap(),
        )
        .unwrap();
    }

    assert_eq!(before, input_snapshot(&input_root));

    if should_cleanup {
        fs::remove_dir_all(probe_root).unwrap();
    }
}

fn input_snapshot(root: &Path) -> Vec<(PathBuf, Vec<u8>)> {
    fn visit(root: &Path, current: &Path, files: &mut Vec<(PathBuf, Vec<u8>)>) {
        for entry in fs::read_dir(current).unwrap() {
            let path = entry.unwrap().path();
            if path.is_dir() {
                visit(root, &path, files);
            } else {
                files.push((
                    path.strip_prefix(root).unwrap().to_path_buf(),
                    fs::read(path).unwrap(),
                ));
            }
        }
    }

    let mut files = Vec::new();
    visit(root, root, &mut files);
    files.sort_by(|left, right| left.0.cmp(&right.0));
    files
}

fn unique_test_dir() -> PathBuf {
    let stamp = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .unwrap()
        .as_nanos();
    env::temp_dir().join(format!("s2-service-{}-{stamp}", std::process::id()))
}
