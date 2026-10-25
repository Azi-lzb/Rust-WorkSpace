#!/usr/bin/env python3
"""Compare the Rust S5 rule-evaluation service with the Python AuditEngine.

Samples enabled 单位贷款 rules of the real configuration, evaluates each rule
on both sides against the same real audit-copy and configuration in one
temporary directory, and diffs the full rule outcome: status, triggered, error
classification, description, current/comparison/difference values+units,
variables, evidence, and retrieval traces. Aggregate counts only.
"""

from __future__ import annotations

import argparse
import dataclasses
import json
import os
import shutil
import subprocess
import sys
import tempfile
from decimal import Decimal
from pathlib import Path

ROOT = Path(__file__).resolve().parents[3]


def normalize_scalar(value, numeric_mode="excel_binary64"):
    if isinstance(value, bool):
        return value
    if isinstance(value, (Decimal, float, int)):
        if numeric_mode == "excel_binary64":
            number = float(value)
            if number.is_integer() and abs(number) < 1e16:
                return f"{number:.1f}"
            return repr(number)
        return str(value)
    return value


def normalize_evidence(items):
    normalized = []
    for item in items:
        record = dataclasses.asdict(item) if dataclasses.is_dataclass(item) else dict(item)
        record.pop("raw_type", None)
        if "raw_value" in record:
            record["raw_value"] = normalize_scalar(record["raw_value"])
        if record.get("mapping_cell") is None:
            record["mapping_cell"] = ""
        if record.get("report_item") is None:
            record["report_item"] = ""
        if record.get("metric_header") is None:
            record["metric_header"] = ""
        if record.get("metric_definition_evidence") is None:
            record["metric_definition_evidence"] = ""
        normalized.append(record)
    return normalized


def normalize_trace(rows):
    normalized = []
    for row in rows:
        row = dict(row)
        for key in ("raw_value", "resolved_value"):
            if isinstance(row.get(key), (Decimal, float, int)) and not isinstance(
                row.get(key), bool
            ):
                row[key] = normalize_scalar(row[key])
        if row.get("row_path") is None:
            row["row_path"] = ""
        if row.get("column_path") is None:
            row["column_path"] = ""
        normalized.append(row)
    return normalized


def rust_scalar(value):
    """Normalize a Rust outcome scalar to the python baseline form."""
    if isinstance(value, bool) or value is None:
        return value
    if isinstance(value, (int, float)):
        return normalize_scalar(value)
    if isinstance(value, str):
        try:
            number = float(value)
        except ValueError:
            return value
        return normalize_scalar(number)
    return value

def python_outcomes(config_copy: Path, source_copy: Path, rule_ids: list[str]):
    from s1audit_configured.engine import AuditEngine
    from s1audit_configured.model import RuntimeContext, SourceBinding
    from s1audit_configured.sources import SourceRepository
    from s1audit_configured.workbook_config import load_configuration

    configuration = load_configuration(config_copy, include_disabled=False, business="")
    engine = AuditEngine(configuration)
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
    repository = SourceRepository(configuration, context)
    outcomes = []
    for rule in configuration.rules:
        if rule.rule_id not in rule_ids:
            continue
        result = engine._evaluate_rule(rule, context, repository, "PARITY")
        outcomes.append({
            "rule_id": rule.rule_id,
            "status": result.status,
            "triggered": result.triggered,
            "error_code": result.error_code,
            "error_message": result.error_message,
            "description": result.description,
            "current_value": normalize_scalar(result.current_value)
            if result.current_value is not None
            else None,
            "comparison_value": normalize_scalar(result.comparison_value)
            if result.comparison_value is not None
            else None,
            "difference_value": normalize_scalar(result.difference_value)
            if result.difference_value is not None
            else None,
            "current_unit": result.current_unit or "",
            "comparison_unit": result.comparison_unit or "",
            "difference_unit": result.difference_unit or "",
            "current_precision": result.current_precision,
            "comparison_precision": result.comparison_precision,
            "difference_precision": result.difference_precision,
            "variables": {
                name: normalize_scalar(value)
                for name, value in result.variables.items()
            },
            "evidence": normalize_evidence(result.evidence),
            "retrieval_trace": normalize_trace(result.retrieval_trace),
            "unused_variables": list(result.unused_variables),
        })
    repository.close()
    return outcomes


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


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--source-root", type=Path, required=True)
    parser.add_argument("--config-workbook", type=Path, required=True)
    parser.add_argument("--audit-copy", type=Path, required=True)
    parser.add_argument("--limit", type=int, default=6)
    parser.add_argument("--workspace", type=Path, default=ROOT)
    parser.add_argument("--report", type=Path, required=True)
    args = parser.parse_args()

    source_root = args.source_root.resolve(strict=True)
    workspace = args.workspace.resolve(strict=True)
    report_path = args.report.resolve()
    sys.path.insert(0, str(source_root / "core" / "src"))
    from s1audit_configured.workbook_config import load_configuration

    with tempfile.TemporaryDirectory(prefix="audit-s5-rule-") as temporary:
        directory = Path(temporary)
        config_copy = directory / "config.xlsx"
        source_copy = directory / "source.xlsx"
        shutil.copy2(args.config_workbook.resolve(strict=True), config_copy)
        shutil.copy2(args.audit_copy.resolve(strict=True), source_copy)

        configuration = load_configuration(config_copy, include_disabled=False, business="")
        rule_ids = [
            rule.rule_id
            for rule in configuration.rules
            if rule.business == "单位贷款" and rule.enabled
        ][: args.limit]

        expected = python_outcomes(config_copy, source_copy, rule_ids)

        request = {
            "id": 1,
            "config": str(config_copy),
            "source": str(source_copy),
            "period": "2026-07-31",
            "institution_code": "PARITY",
            "institution_name": "PARITY",
            "source_id": "CURRENT",
            "rule_ids": rule_ids,
        }
        input_path = directory / "request.json"
        input_path.write_text(json.dumps(request, ensure_ascii=False), encoding="utf-8")
        rust_report = directory / "rust_outcomes.jsonl"
        environment = os.environ.copy()
        environment["AUDIT_S5_RULE_INPUT"] = str(input_path)
        environment["AUDIT_S5_RULE_REPORT"] = str(rust_report)
        completed = subprocess.run(
            [
                "cargo", "test", "-p", "audit-core", "--test", "s5_rule_parity",
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
        payload = json.loads(rust_report.read_text(encoding="utf-8").splitlines()[0])
        rust_by_rule = {
            entry["rule_id"]: entry["outcome"]
            for entry in payload.get("results", [])
        }

        matched = 0
        mismatched_rules: list[str] = []
        status_distribution: dict[str, int] = {}
        for expected_outcome in expected:
            rule_id = expected_outcome["rule_id"]
            actual = rust_by_rule.get(rule_id, {})
            status_distribution[expected_outcome["status"]] = (
                status_distribution.get(expected_outcome["status"], 0) + 1
            )
            actual_normalized = {
                "status": actual.get("status"),
                "triggered": actual.get("triggered"),
                "error_code": actual.get("error_code") or "",
                "error_message": actual.get("error_message") or "",
                "description": actual.get("description") or "",
                "current_value": rust_scalar(actual.get("current_value")),
                "comparison_value": rust_scalar(actual.get("comparison_value")),
                "difference_value": rust_scalar(actual.get("difference_value")),
                "current_unit": actual.get("current_unit") or "",
                "comparison_unit": actual.get("comparison_unit") or "",
                "difference_unit": actual.get("difference_unit") or "",
                "current_precision": actual.get("current_precision"),
                "comparison_precision": actual.get("comparison_precision"),
                "difference_precision": actual.get("difference_precision"),
                "variables": {
                    name: rust_scalar(value)
                    for name, value in (actual.get("variables") or {}).items()
                },
                "evidence": normalize_evidence(actual.get("evidence", [])),
                "retrieval_trace": normalize_trace(actual.get("retrieval_trace", [])),
                "unused_variables": actual.get("unused_variables") or [],
            }
            diffs: list[str] = []
            expected_compare = {
                key: value for key, value in expected_outcome.items() if key != "rule_id"
            }
            diff(expected_compare, actual_normalized, rule_id, diffs)
            if diffs:
                mismatched_rules.append(rule_id)
                if len(mismatched_rules) <= 2:
                    print(json.dumps({
                        "dbg_rule": rule_id,
                        "diffs": diffs[:10],
                    }, ensure_ascii=False))
            else:
                matched += 1

    report = {
        "schema_version": 1,
        "suite": "audit.s5_rule_evaluation_parity",
        "status": "completed" if not mismatched_rules else "mismatch",
        "sample_scope": {
            "business": "单位贷款",
            "rules_sampled": len(expected),
            "files_copied_to_temporary": True,
        },
        "comparison": {
            "matched": matched,
            "mismatch_rules": mismatched_rules[:20],
            "mismatch_count": len(mismatched_rules),
            "python_status_distribution": dict(sorted(status_distribution.items())),
        },
        "verification": {
            "rust_runner": "audit-core/tests/s5_rule_parity.rs",
            "python_baseline": "s1audit_configured.engine.AuditEngine._evaluate_rule",
            "inputs_unchanged": True,
            "raw_values_persisted": False,
            "capability_status": "unverified",
            "production_route_eligible": False,
        },
        "limitations": [
            "No applicability-scope (@当前行) expansion: real configuration rules "
            "without scopes are covered; scope expansion belongs to the batch "
            "orchestration slice.",
            "Retrieval is the single-binding current-source path without "
            "contracts, matching the real configuration shape.",
        ],
    }
    report_path.parent.mkdir(parents=True, exist_ok=True)
    report_path.write_text(
        json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )
    print(json.dumps({
        "status": report["status"],
        "rules": len(expected),
        "matched": matched,
        "mismatches": len(mismatched_rules),
    }))
    return 0 if not mismatched_rules else 1


if __name__ == "__main__":
    raise SystemExit(main())
