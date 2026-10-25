#!/usr/bin/env python3
"""Compare the Rust S5 run orchestration with the Python AuditEngine.run.

Runs both sides over the full enabled 单位贷款 rule set against the same real
configuration and audit copy (one temporary directory), then diffs every rule
result row plus the status counts and run status. Aggregate counts only.
"""

from __future__ import annotations

import argparse
import json
import os
import shutil
import subprocess
import sys
import tempfile
from decimal import Decimal
from pathlib import Path

ROOT = Path(__file__).resolve().parents[3]


def normalize_scalar(value):
    if isinstance(value, bool) or value is None:
        return value
    if isinstance(value, (Decimal, float, int)):
        number = float(value)
        if number.is_integer() and abs(number) < 1e16:
            return f"{number:.1f}"
        return repr(number)
    return value


def rust_scalar(value):
    if isinstance(value, bool) or value is None:
        return value
    if isinstance(value, (int, float)):
        return normalize_scalar(value)
    if isinstance(value, str):
        try:
            return normalize_scalar(float(value))
        except ValueError:
            return value
    return value


def python_run(
    config_copy: Path,
    source_copy: Path,
    snapshot_dir: Path,
    template_path: Path | None,
    template_dir: Path | None,
):
    from s1audit_configured.engine import AuditEngine
    from s1audit_configured.model import RuntimeContext, SourceBinding
    from s1audit_configured.runtime import RuntimePaths

    engine = AuditEngine(config_copy)
    context = RuntimeContext(
        period="2026-07-31",
        institution_code="PARITY",
        institution_name="PARITY",
        current_sources={
            "CURRENT": SourceBinding(
                "CURRENT",
                source_copy,
                business="单位贷款",
                period="2026-07-31",
                institution_code="PARITY",
            )
        },
    )
    result = engine.run(
        context,
        businesses=["单位贷款"],
        schema_snapshot_dir=snapshot_dir,
        persist_schema_snapshots=False,
        runtime_paths=RuntimePaths(
            template_path=template_path,
            template_dir=template_dir,
        ),
    )
    return result


def diff(left, right, path: str, diffs: list[str]) -> None:
    if isinstance(left, dict) and isinstance(right, dict):
        for key in sorted(set(left) | set(right)):
            diff(left.get(key), right.get(key), f"{path}.{key}", diffs)
    elif isinstance(left, list) and isinstance(right, list):
        if len(left) != len(right):
            diffs.append(f"{path}: length {len(left)} vs {len(right)}")
        for index in range(min(len(left), len(right))):
            diff(left[index], right[index], f"{path}[{index}]", diffs)
    elif left != right:
        diffs.append(path)


COMPARE_FIELDS = (
    "rule_id",
    "revision",
    "rule_name",
    "business",
    "period",
    "scope_path",
    "status",
    "triggered",
    "severity",
    "check_type",
    "metric",
    "description",
    "sheet_name",
    "formula_cell",
    "current_value",
    "current_unit",
    "current_precision",
    "comparison_value",
    "comparison_unit",
    "comparison_precision",
    "difference_value",
    "difference_unit",
    "difference_precision",
    "error_code",
    "error_message",
)


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--source-root", type=Path, required=True)
    parser.add_argument("--config-workbook", type=Path, required=True)
    parser.add_argument("--audit-copy", type=Path, required=True)
    parser.add_argument("--template-path", type=Path)
    parser.add_argument("--template-dir", type=Path)
    parser.add_argument("--workspace", type=Path, default=ROOT)
    parser.add_argument("--report", type=Path, required=True)
    args = parser.parse_args()

    source_root = args.source_root.resolve(strict=True)
    workspace = args.workspace.resolve(strict=True)
    report_path = args.report.resolve()
    template_path = args.template_path.resolve(strict=True) if args.template_path else None
    template_dir = args.template_dir.resolve(strict=True) if args.template_dir else None
    sys.path.insert(0, str(source_root / "core" / "src"))

    with tempfile.TemporaryDirectory(prefix="audit-s5-run-") as temporary:
        directory = Path(temporary)
        config_copy = directory / "config.xlsx"
        source_copy = directory / "source.xlsx"
        shutil.copy2(args.config_workbook.resolve(strict=True), config_copy)
        shutil.copy2(args.audit_copy.resolve(strict=True), source_copy)

        run_result = python_run(
            config_copy,
            source_copy,
            directory / "snap",
            template_path,
            template_dir,
        )
        expected_rows = [
            {field: row.get(field) for field in COMPARE_FIELDS}
            for row in run_result.ledger["rules"]
        ]
        expected_counts = dict(run_result.ledger["status_counts"])

        request = {
            "config": str(config_copy),
            "source": str(source_copy),
            "period": "2026-07-31",
            "institution_code": "PARITY",
            "institution_name": "PARITY",
            "source_id": "CURRENT",
            "businesses": ["单位贷款"],
            "template_path": str(template_path) if template_path else None,
            "template_dir": str(template_dir) if template_dir else None,
        }
        input_path = directory / "request.json"
        input_path.write_text(json.dumps(request, ensure_ascii=False), encoding="utf-8")
        rust_report = directory / "rust_run.json"
        environment = os.environ.copy()
        environment["AUDIT_S5_RUN_INPUT"] = str(input_path)
        environment["AUDIT_S5_RUN_REPORT"] = str(rust_report)
        completed = subprocess.run(
            [
                "cargo", "test", "-p", "audit-core", "--test", "s5_run_parity",
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
                "exit_code": completed.returncode,
                "stderr_tail": completed.stderr[-1500:],
            }))
            return 1
        payload = json.loads(rust_report.read_text(encoding="utf-8"))
        rust_rows = payload.get("results", [])
        rust_counts = payload.get("status_counts", {})

        matched = 0
        mismatched: list[int] = []
        count_mismatch = expected_counts != rust_counts
        for index, expected in enumerate(expected_rows):
            actual_raw = rust_rows[index] if index < len(rust_rows) else {}
            actual = {}
            for field in COMPARE_FIELDS:
                value = actual_raw.get(field)
                if field in (
                    "current_value",
                    "comparison_value",
                    "difference_value",
                ):
                    value = rust_scalar(value)
                if field in ("error_code", "error_message", "description", "scope_path",
                             "current_unit", "comparison_unit", "difference_unit",
                             "sheet_name", "formula_cell"):
                    value = value or ""
                if field == "triggered" and value is None and expected.get(field) is None:
                    value = None
                actual[field] = value
            expected_normalized = dict(expected)
            for field in ("current_value", "comparison_value", "difference_value"):
                expected_normalized[field] = normalize_scalar(expected.get(field))
            for field in ("error_code", "error_message", "scope_path", "current_unit",
                          "comparison_unit", "difference_unit"):
                expected_normalized[field] = expected.get(field) or ""
            diffs: list[str] = []
            diff(expected_normalized, actual, f"row[{index}]", diffs)
            if diffs:
                mismatched.append(index)
                if len(mismatched) <= 3:
                    print(json.dumps({
                        "row_index": index,
                        "mismatch_fields": diffs[:10],
                    }, ensure_ascii=False))
            else:
                matched += 1

    report = {
        "schema_version": 1,
        "suite": "audit.s5_run_orchestration_parity",
        "status": "completed" if not mismatched and not count_mismatch else "mismatch",
        "sample_scope": {
            "business": "单位贷款",
            "rules_executed": len(expected_rows),
            "files_copied_to_temporary": True,
            "template_path_supplied": template_path is not None,
            "template_dir_supplied": template_dir is not None,
        },
        "comparison": {
            "rows_matched": matched,
            "rows_mismatched": len(mismatched),
            "mismatch_row_indices": mismatched[:10],
            "status_counts_match": not count_mismatch,
            "python_status_counts": expected_counts,
            "rust_status_counts": rust_counts,
            "python_run_status": run_result.status,
        },
        "verification": {
            "rust_runner": "audit-core/tests/s5_run_parity.rs",
            "python_baseline": "s1audit_configured.engine.AuditEngine.run",
            "python_baseline_version": sys.version.split()[0],
            "inputs_unchanged": True,
            "raw_values_persisted": False,
            "capability_status": "unverified",
            "production_route_eligible": False,
        },
        "limitations": [
            "Hash-integrity downgrade, schema snapshots, data-quality warnings, "
            "output writing, and V_PRE/V_EXT bindings are not exercised; the run "
            "covers the single-binding current-source path with unchanged inputs.",
        ],
    }
    report_path.parent.mkdir(parents=True, exist_ok=True)
    report_path.write_text(
        json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )
    print(json.dumps({
        "status": report["status"],
        "rows": len(expected_rows),
        "matched": matched,
        "mismatches": len(mismatched),
        "counts_match": not count_mismatch,
    }))
    return 0 if not mismatched and not count_mismatch else 1


if __name__ == "__main__":
    raise SystemExit(main())
