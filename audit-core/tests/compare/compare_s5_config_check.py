#!/usr/bin/env python3
"""Compare the Rust S5 config-check port with the Python baseline.

Runs `validate_s5_config` on both sides for a set of scopes (configuration
level only: no template or external file, which selects the ported
not-checked branches), then diffs the full result dictionaries field by
field. Only aggregate counts are persisted.
"""

from __future__ import annotations

import argparse
import json
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[3]

CASES = (
    {"scope": "current"},
    {"scope": "current", "business": "单位贷款"},
    {"scope": "current_business"},
)


def diff_values(left, right, path: str, diffs: list[str]) -> None:
    if isinstance(left, dict) and isinstance(right, dict):
        for key in sorted(set(left) | set(right)):
            if key not in left:
                diffs.append(f"{path}.{key}: missing_in_python")
            elif key not in right:
                diffs.append(f"{path}.{key}: missing_in_rust")
            else:
                diff_values(left[key], right[key], f"{path}.{key}", diffs)
    elif isinstance(left, list) and isinstance(right, list):
        if len(left) != len(right):
            diffs.append(f"{path}: length {len(left)} vs {len(right)}")
        for index in range(min(len(left), len(right))):
            diff_values(left[index], right[index], f"{path}[{index}]", diffs)
    elif left != right:
        diffs.append(path)


def run_case(case: dict, python_side) -> dict:
    try:
        report = python_side(
            config_path=case["config"],
            template_path=case.get("template_path", ""),
            external_path=case.get("external_path", ""),
            business=case.get("business", ""),
            scope=case["scope"],
        )
        return {"status": "ok", "report": json.loads(json.dumps(report, ensure_ascii=False))}
    except Exception as error:
        return {
            "status": "error",
            "error_kind": type(error).__name__,
            "message": str(error),
        }


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--source-root", type=Path, required=True)
    parser.add_argument("--config-workbook", type=Path, required=True)
    parser.add_argument("--workspace", type=Path, default=ROOT)
    parser.add_argument("--report", type=Path, required=True)
    args = parser.parse_args()

    source_root = args.source_root.resolve(strict=True)
    workspace = args.workspace.resolve(strict=True)
    report_path = args.report.resolve()
    sys.path.insert(0, str(source_root / "core" / "src"))
    from base_audit.systems.s5_configured_statistics.service import (
        validate_s5_config as python_validate_s5_config,
    )

    with tempfile.TemporaryDirectory(prefix="audit-s5-config-check-") as temporary:
        directory = Path(temporary)
        config_copy = directory / "config.xlsx"
        shutil.copy2(args.config_workbook.resolve(strict=True), config_copy)

        case_reports = []
        total_diffs = 0
        for case_index, base_case in enumerate(CASES):
            case = {**base_case, "config": str(config_copy)}
            python_payload = run_case(case, python_validate_s5_config)

            request = {
                "config": case["config"],
                "template_path": case.get("template_path", ""),
                "external_path": case.get("external_path", ""),
                "business": case.get("business", ""),
                "scope": case["scope"],
            }
            request_path = directory / f"request_{case_index}.json"
            request_path.write_text(
                json.dumps(request, ensure_ascii=False), encoding="utf-8"
            )
            rust_report = directory / f"rust_{case_index}.json"
            environment = os.environ.copy()
            environment["AUDIT_S5_CHECK_INPUT"] = str(request_path)
            environment["AUDIT_S5_CHECK_REPORT"] = str(rust_report)
            completed = subprocess.run(
                [
                    "cargo", "test", "-p", "audit-core", "--test", "s5_config_check_parity",
                    "--locked", "--", "--ignored", "--nocapture",
                ],
                cwd=workspace,
                env=environment,
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                check=False,
            )
            if completed.returncode != 0 or not rust_report.is_file():
                print(json.dumps({
                    "status": "rust_runner_failed",
                    "case": case["scope"],
                    "exit_code": completed.returncode,
                    "stderr_tail": completed.stderr[-1500:],
                }))
                return 1
            rust_payload = json.loads(rust_report.read_text(encoding="utf-8"))

            diffs: list[str] = []
            if python_payload["status"] != rust_payload.get("status"):
                diffs.append("status")
            elif python_payload["status"] == "ok":
                python_report = dict(python_payload["report"])
                python_report.pop("path", None)
                rust_report_dict = dict(rust_payload.get("report", {}))
                diff_values(python_report, rust_report_dict, "report", diffs)
            else:
                if python_payload.get("message") != rust_payload.get("message"):
                    diffs.append("message")
            total_diffs += len(diffs)
            case_reports.append({
                "scope": case["scope"],
                "business": case.get("business", ""),
                "python_status": python_payload["status"],
                "rust_status": rust_payload.get("status"),
                "diff_count": len(diffs),
                "diff_paths_sample": sorted(set(diffs))[:15],
                "python_valid": python_payload.get("report", {}).get("valid")
                if python_payload["status"] == "ok"
                else None,
                "python_complete": python_payload.get("report", {}).get("complete")
                if python_payload["status"] == "ok"
                else None,
            })

        report = {
            "schema_version": 1,
            "suite": "audit.s5_config_check_parity",
            "status": "completed" if total_diffs == 0 else "mismatch",
            "sample_scope": {
                "config_level_only": True,
                "cases": [
                    {"scope": case["scope"], "business": case.get("business", "")}
                    for case in CASES
                ],
                "config_copied_to_temporary": True,
            },
            "comparison": {
                "case_count": len(CASES),
                "total_diff_count": total_diffs,
                "cases": case_reports,
            },
            "verification": {
                "rust_runner": "audit-core/tests/s5_config_check_parity.rs",
                "python_baseline": (
                    "base_audit.systems.s5_configured_statistics."
                    "service.validate_s5_config"
                ),
                "inputs_unchanged": True,
                "raw_paths_or_messages_persisted": False,
                "capability_status": "unverified",
                "production_route_eligible": False,
            },
            "limitations": [
                "Template/external workbook address scanning (including the "
                "all/all_enabled full-scope branch) is not migrated; providing "
                "a template or external file selects the not-migrated boundary "
                "on the Rust side and is excluded from parity.",
                "Load-error variants attach empty locations on the Rust side "
                "(Python row.check adds cell locations); the real configuration "
                "loads cleanly so no error variant is exercised.",
            ],
        }
        report_path.parent.mkdir(parents=True, exist_ok=True)
        report_path.write_text(
            json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
        )
        print(json.dumps({
            "status": report["status"],
            "cases": len(CASES),
            "total_diffs": total_diffs,
        }))
        return 0 if total_diffs == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())
