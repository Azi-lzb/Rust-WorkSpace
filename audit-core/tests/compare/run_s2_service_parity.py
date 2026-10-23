"""Run the read-only Python S2 service and compare it with the Rust service candidate."""

from __future__ import annotations

import argparse
import dataclasses
import hashlib
import importlib.util
import json
import os
import re
import shlex
import shutil
import subprocess
import sys
import tempfile
from decimal import Decimal
from pathlib import Path
from typing import Any


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def plain(value: Any) -> Any:
    if dataclasses.is_dataclass(value):
        return {field.name: plain(getattr(value, field.name)) for field in dataclasses.fields(value)}
    if isinstance(value, dict):
        return {key: plain(item) for key, item in value.items()}
    if isinstance(value, (tuple, list)):
        return [plain(item) for item in value]
    if isinstance(value, Path):
        return str(value)
    return value


def numeric(value: Any) -> bool:
    return isinstance(value, (int, float, Decimal)) and not isinstance(value, bool)


def compare_tree(expected: Any, actual: Any, path: str, differences: list[dict[str, Any]]) -> None:
    if type(expected) is int and type(actual) is int:
        if expected != actual:
            differences.append({"path": path, "expected": expected, "actual": actual})
        return
    if numeric(expected) and numeric(actual):
        left, right = Decimal(str(expected)), Decimal(str(actual))
        if left == right:
            return
        scale = max(abs(left), abs(right))
        relative = abs(left - right) / scale if scale else Decimal(0)
        if scale and relative <= Decimal("1e-15"):
            return
        differences.append({"path": path, "expected": expected, "actual": actual})
        return
    if isinstance(expected, dict) and isinstance(actual, dict):
        expected_keys, actual_keys = set(expected), set(actual)
        if expected_keys != actual_keys:
            differences.append({
                "path": path,
                "missing_in_actual": sorted(expected_keys - actual_keys),
                "unexpected_in_actual": sorted(actual_keys - expected_keys),
            })
        for key in sorted(expected_keys & actual_keys):
            compare_tree(expected[key], actual[key], f"{path}.{key}", differences)
        return
    if isinstance(expected, list) and isinstance(actual, list):
        if len(expected) != len(actual):
            differences.append({"path": path, "expected_length": len(expected), "actual_length": len(actual)})
        for index, (left, right) in enumerate(zip(expected, actual)):
            compare_tree(left, right, f"{path}[{index}]", differences)
        return
    if isinstance(expected, str) and isinstance(actual, str):
        if normalize_path_text(expected) == normalize_path_text(actual):
            return
        differences.append({"path": path, "expected": expected, "actual": actual})
        return
    if type(expected) is not type(actual) or expected != actual:
        differences.append({"path": path, "expected": expected, "actual": actual})


def normalize_path_text(value: str) -> str:
    if not re.search(r"(?i)(?:[a-z]:[\\/]|/mnt/[a-z]/)", value):
        return value
    normalized = value.replace("\\", "/")
    return re.sub(
        r"(?i)([a-z]):/",
        lambda match: f"/mnt/{match.group(1).lower()}/",
        normalized,
    )


def path_from_wsl(value: str) -> Path:
    match = re.match(r"^/mnt/([a-z])/(.*)$", value)
    if os.name == "nt" and match:
        return Path(f"{match.group(1).upper()}:/{match.group(2)}")
    return Path(value)


def progress_for_comparison(events: list[dict[str, Any]]) -> list[dict[str, Any]]:
    normalized = []
    for event in events:
        item = dict(event)
        if item.get("stage") == "external_read" and item.get("percent") == 76:
            # Engine/version and elapsed time are intentionally reader-specific.
            item["message"] = "<external-reader-result>"
        else:
            item["message"] = normalize_path_text(str(item.get("message", "")))
        normalized.append(item)
    return normalized


def run_rust_test(workspace: Path, probe_root: Path, distro: str) -> None:
    if os.name == "nt" and shutil.which("wsl.exe"):
        def wslpath(path: Path) -> str:
            result = subprocess.run(
                ["wsl.exe", "-d", distro, "--", "wslpath", "-a", str(path).replace("\\", "/")],
                check=True,
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
            )
            return result.stdout.strip()

        workspace_wsl = wslpath(workspace)
        probe_wsl = wslpath(probe_root)
        command = (
            f"cd {shlex.quote(workspace_wsl)} && "
            f"AUDIT_S2_SERVICE_PROBE_DIR={shlex.quote(probe_wsl)} "
            "cargo test --locked -p audit-core --test s2_service_parity -- --nocapture"
        )
        subprocess.run(["wsl.exe", "-d", distro, "--", "bash", "-lc", command], check=True)
        return

    environment = os.environ.copy()
    environment["AUDIT_S2_SERVICE_PROBE_DIR"] = str(probe_root)
    subprocess.run(
        ["cargo", "test", "--locked", "-p", "audit-core", "--test", "s2_service_parity", "--", "--nocapture"],
        cwd=workspace,
        env=environment,
        check=True,
    )


def python_case(source_root: Path, input_root: Path, output_root: Path, scenario: str):
    sys.path.insert(0, str(source_root / "core" / "src"))
    from base_audit.systems.s2_report_collection.service import PeriodAuditV2Service

    events: list[dict[str, Any]] = []
    central_path = input_root / "plain_layout.xlsx" if scenario == "with_external" else None
    config_path = input_root / (
        "check_gap_duplicate.xlsx" if scenario == "invalid_config" else "check_valid.xlsx"
    )
    try:
        result = PeriodAuditV2Service().run(
            current_dir=input_root / "current",
            previous_dir=input_root / "previous",
            config_path=config_path,
            output_dir=output_root / "output",
            central_path=central_path,
            on_progress=events.append,
        )
    except Exception as error:
        if scenario != "invalid_config":
            raise
        return {"error": str(error), "progress": plain(events), "workbook": None}
    return {
        "findings": plain(result.findings),
        "summary": plain(result.summary),
        "progress": plain(events),
        "workbook": Path(result.paths.result),
    }


def load_workbook_comparator(audit_root: Path):
    path = audit_root / "tests" / "compare" / "run_s2_exporter_parity.py"
    spec = importlib.util.spec_from_file_location("s2_exporter_parity", path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"cannot load workbook comparison helper: {path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module.compare_workbooks


def json_default(value: Any):
    if isinstance(value, Decimal):
        return str(value)
    if isinstance(value, Path):
        return str(value)
    raise TypeError(f"not JSON serializable: {type(value).__name__}")


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--source-root", type=Path, required=True)
    parser.add_argument("--report-dir", type=Path)
    parser.add_argument("--wsl-distro", default="Ubuntu")
    args = parser.parse_args()

    source_root = args.source_root.resolve()
    audit_root = Path(__file__).resolve().parents[2]
    workspace = audit_root.parent
    fixture_root = audit_root / "tests" / "fixtures"
    report_dir = args.report_dir or audit_root / "probe" / "reports"
    report_dir.mkdir(parents=True, exist_ok=True)
    source_files = [
        "service.py", "config.py", "config_check.py", "importer.py", "validation.py",
        "period_engine.py", "external_engine.py", "rule_engine.py", "exporter.py", "models.py",
        "unit_conversion.py",
    ]
    source_hashes_before = {
        name: sha256(source_root / "core/src/base_audit/systems/s2_report_collection" / name)
        for name in source_files
    }
    workbook_compare = load_workbook_comparator(audit_root)
    report: dict[str, Any] = {
        "status": "failed",
        "baseline_python": sys.version.split()[0],
        "source_root": str(source_root),
        "source_sha256": source_hashes_before,
        "numeric_relative_tolerance": "1e-15",
        "scenarios": {},
    }

    with tempfile.TemporaryDirectory(prefix="s2-service-parity-") as temp_name:
        probe_root = Path(temp_name)
        run_rust_test(workspace, probe_root, args.wsl_distro)
        input_root = probe_root / "input"
        inputs = [
            input_root / "current" / "9144#2026-06-30#x#20202#甲银行.xlsx",
            input_root / "previous" / "9144#2026-03-31#x#20202#甲银行.xlsx",
            input_root / "check_valid.xlsx",
            input_root / "check_gap_duplicate.xlsx",
            input_root / "plain_layout.xlsx",
        ]
        hashes_before_python = {path.relative_to(input_root).as_posix(): sha256(path) for path in inputs}
        for scenario in ("without_external", "with_external", "invalid_config"):
            python_root = probe_root / "python" / scenario
            python_result = python_case(source_root, input_root, python_root, scenario)
            rust_root = probe_root / "rust" / scenario
            rust_progress_path = rust_root / "progress.json"
            rust_progress = json.loads(rust_progress_path.read_text(encoding="utf-8"), parse_float=Decimal)
            rust_cli_root = probe_root / "rust_cli" / scenario
            rust_cli_result = json.loads(
                (rust_cli_root / "cli_result.json").read_text(encoding="utf-8"),
                parse_float=Decimal,
            )
            rust_cli_progress = json.loads(
                (rust_cli_root / "progress.json").read_text(encoding="utf-8"),
                parse_float=Decimal,
            )

            business_differences: list[dict[str, Any]] = []
            workbook_stats: dict[str, Any] = {}
            workbook_differences: list[dict[str, Any]] = []
            path_differences: list[dict[str, Any]] = []
            rust_workbook_path = None
            python_output_created = (python_root / "output").exists()
            rust_output_created = (rust_root / "output").exists()
            cli_business_differences: list[dict[str, Any]] = []
            cli_workbook_stats: dict[str, Any] = {}
            cli_workbook_differences: list[dict[str, Any]] = []
            cli_path_differences: list[dict[str, Any]] = []
            cli_report_path = None
            cli_output_created = (rust_cli_root / "output").exists()
            if scenario == "invalid_config":
                rust_error = (rust_root / "error.txt").read_text(encoding="utf-8")
                compare_tree(python_result["error"], rust_error, "error", business_differences)
                compare_tree(
                    python_result["error"],
                    rust_cli_result.get("error"),
                    "cli.error",
                    cli_business_differences,
                )
            else:
                rust_analysis = json.loads(
                    (rust_root / "analysis.json").read_text(encoding="utf-8"),
                    parse_float=Decimal,
                )
                rust_workbook_path = path_from_wsl(
                    (rust_root / "report_path.txt").read_text(encoding="utf-8").strip()
                )
                compare_tree(python_result["findings"], rust_analysis["findings"], "findings", business_differences)
                compare_tree(python_result["summary"], rust_analysis["summary"], "summary", business_differences)
                workbook_stats, raw_workbook_differences = workbook_compare(
                    python_result["workbook"], rust_workbook_path, 1e-15
                )
                path_differences = [
                    difference
                    for difference in raw_workbook_differences
                    if difference.get("kind") == "value"
                    and isinstance(difference.get("python"), str)
                    and isinstance(difference.get("rust"), str)
                    and normalize_path_text(difference["python"]) == normalize_path_text(difference["rust"])
                ]
                workbook_differences = [
                    difference
                    for difference in raw_workbook_differences
                    if difference not in path_differences
                ]
                if rust_cli_result.get("status") != "completed":
                    cli_business_differences.append({
                        "path": "cli.status",
                        "expected": "completed",
                        "actual": rust_cli_result.get("status"),
                    })
                compare_tree(
                    python_result["summary"],
                    rust_cli_result.get("summary"),
                    "cli.summary",
                    cli_business_differences,
                )
                cli_report_path = path_from_wsl(rust_cli_result["report_path"])
                cli_workbook_stats, raw_cli_workbook_differences = workbook_compare(
                    python_result["workbook"], cli_report_path, 1e-15
                )
                cli_path_differences = [
                    difference
                    for difference in raw_cli_workbook_differences
                    if difference.get("kind") == "value"
                    and isinstance(difference.get("python"), str)
                    and isinstance(difference.get("rust"), str)
                    and normalize_path_text(difference["python"]) == normalize_path_text(difference["rust"])
                ]
                cli_workbook_differences = [
                    difference
                    for difference in raw_cli_workbook_differences
                    if difference not in cli_path_differences
                ]
            progress_differences: list[dict[str, Any]] = []
            compare_tree(
                progress_for_comparison(python_result["progress"]),
                progress_for_comparison(rust_progress),
                "progress",
                progress_differences,
            )
            cli_progress_differences: list[dict[str, Any]] = []
            compare_tree(
                progress_for_comparison(python_result["progress"]),
                progress_for_comparison(rust_cli_progress),
                "cli.progress",
                cli_progress_differences,
            )
            report["scenarios"][scenario] = {
                "findings": len(python_result.get("findings", [])),
                "summary": plain(python_result.get("summary")),
                "progress_events": len(python_result["progress"]),
                "business_difference_count": len(business_differences),
                "business_differences": business_differences[:100],
                "progress_difference_count": len(progress_differences),
                "progress_differences": progress_differences[:100],
                "workbook_stats": workbook_stats,
                "workbook_difference_count": len(workbook_differences),
                "workbook_differences": workbook_differences[:100],
                "platform_path_difference_count": len(path_differences),
                "rust_workbook_sha256": sha256(rust_workbook_path) if rust_workbook_path else None,
                "python_workbook_sha256": sha256(python_result["workbook"]) if python_result["workbook"] else None,
                "error_message": normalize_path_text(rust_error) if scenario == "invalid_config" else None,
                "python_output_created": python_output_created,
                "rust_output_created": rust_output_created,
                "cli_summary": plain(rust_cli_result.get("summary")),
                "cli_progress_events": len(rust_cli_progress),
                "cli_business_difference_count": len(cli_business_differences),
                "cli_business_differences": cli_business_differences[:100],
                "cli_progress_difference_count": len(cli_progress_differences),
                "cli_progress_differences": cli_progress_differences[:100],
                "cli_workbook_stats": cli_workbook_stats,
                "cli_workbook_difference_count": len(cli_workbook_differences),
                "cli_workbook_differences": cli_workbook_differences[:100],
                "cli_platform_path_difference_count": len(cli_path_differences),
                "cli_report_sha256": sha256(cli_report_path) if cli_report_path else None,
                "cli_output_created": cli_output_created,
                "cli_error_message": normalize_path_text(rust_cli_result["error"])
                if scenario == "invalid_config" else None,
            }
        hashes_after_python = {path.relative_to(input_root).as_posix(): sha256(path) for path in inputs}
        report["input_sha256_before_python"] = hashes_before_python
        report["input_sha256_after_python"] = hashes_after_python
        report["inputs_unchanged"] = hashes_before_python == hashes_after_python
        report["rust_inputs_unchanged"] = True  # asserted by the Rust integration test

    report["source_unchanged"] = source_hashes_before == {
        name: sha256(source_root / "core/src/base_audit/systems/s2_report_collection" / name)
        for name in source_files
    }
    report["status"] = "passed" if report["inputs_unchanged"] and report["source_unchanged"] and all(
        scenario["business_difference_count"] == 0
        and scenario["progress_difference_count"] == 0
        and scenario["workbook_difference_count"] == 0
        and scenario["cli_business_difference_count"] == 0
        and scenario["cli_progress_difference_count"] == 0
        and scenario["cli_workbook_difference_count"] == 0
        and (scenario["python_output_created"] == scenario["rust_output_created"])
        and (scenario["python_output_created"] is False if scenario["error_message"] is not None else True)
        and (scenario["python_output_created"] == scenario["cli_output_created"])
        and (scenario["cli_output_created"] is False if scenario["cli_error_message"] is not None else True)
        for scenario in report["scenarios"].values()
    ) else "failed"
    output = report_dir / "s2_service_parity.json"
    output.write_text(json.dumps(report, ensure_ascii=False, indent=2, default=json_default) + "\n", encoding="utf-8")
    print(json.dumps({
        "status": report["status"],
        "report": str(output),
        "scenarios": {
            name: {
                "findings": item["findings"],
                "summary": item["summary"],
                "business_differences": item["business_difference_count"],
                "progress_differences": item["progress_difference_count"],
                "workbook_differences": item["workbook_difference_count"],
                "cli_business_differences": item["cli_business_difference_count"],
                "cli_progress_differences": item["cli_progress_difference_count"],
                "cli_workbook_differences": item["cli_workbook_difference_count"],
            }
            for name, item in report["scenarios"].items()
        },
        "inputs_unchanged": report["inputs_unchanged"],
        "source_unchanged": report["source_unchanged"],
    }, ensure_ascii=False, indent=2))
    return 0 if report["status"] == "passed" else 1


if __name__ == "__main__":
    raise SystemExit(main())
