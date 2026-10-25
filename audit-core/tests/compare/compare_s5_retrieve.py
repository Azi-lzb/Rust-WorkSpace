#!/usr/bin/env python3
"""Compare the Rust S5 retrieval core with the Python SourceRepository.

Builds a temporary working directory holding the real audit configuration and
one real audit copy, samples enabled business rules' V_CUR dictionary calls,
retrieves each on both sides, and compares value/unit/metric-type/errors and
the evidence records (raw_type dropped: the two readers report cell classes
differently). Aggregate counts only.
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


def normalize_number(value):
    if isinstance(value, bool):
        return value
    if isinstance(value, (Decimal, float, int)):
        number = float(value)
        return f"{number:.1f}" if number.is_integer() and abs(number) < 1e16 else repr(number)
    return value


def normalize_evidence(items):
    normalized = []
    for item in items:
        record = dataclasses.asdict(item) if dataclasses.is_dataclass(item) else dict(item)
        record.pop("raw_type", None)
        if "raw_value" in record:
            record["raw_value"] = normalize_number(record["raw_value"])
        normalized.append(record)
    return normalized


def normalize_value(stored, mode):
    if isinstance(stored, bool):
        return stored
    if isinstance(stored, (Decimal, float, int)):
        if mode == "excel_binary64":
            number = float(stored)
            return f"{number:.1f}" if number.is_integer() and abs(number) < 1e16 else repr(number)
        return str(stored)
    return stored


def python_calls(sample_root: Path, config_copy: Path, audit_copy: Path, limit: int):
    from s1audit_configured.dsl import parse_dictionary, parse_expression, calls_in
    from s1audit_configured.errors import RetrievalError
    from s1audit_configured.model import RuntimeContext, SourceBinding
    from s1audit_configured.sources import SourceRepository
    from s1audit_configured.workbook_config import load_configuration

    configuration = load_configuration(config_copy, include_disabled=False, business="")
    context = RuntimeContext(
        period="2026-07-31",
        institution_code="INST",
        institution_name="PARITY",
        current_sources={
            "CURRENT": SourceBinding(
                "CURRENT",
                audit_copy,
                business="单位贷款",
                period="2026-07-31",
                institution_code="INST",
            )
        },
    )
    repository = SourceRepository(configuration, context)
    sampled = []
    for rule in configuration.rules:
        if rule.business != "单位贷款" or not rule.enabled:
            continue
        try:
            dictionary = parse_dictionary(rule.value_dictionary or None)
        except Exception:
            continue
        found = []
        for node in dictionary.values():
            for call in calls_in(node):
                if call.name == "V_CUR":
                    args = tuple(
                        tuple(argument.values)
                        if hasattr(argument, "values") and not hasattr(argument, "kind")
                        else argument.value
                        for argument in call.arguments
                    )
                    found.append((args, dict(call.keywords)))
        for args, keywords in found[: max(1, limit // 8)]:
            sampled.append((rule, args, keywords))
        if len(sampled) >= limit:
            break

    outcomes = []
    for rule, args, keywords in sampled:
        try:
            value = repository.retrieve("V_CUR", args, keywords, rule)
            outcomes.append({
                "status": "ok",
                "result": {
                    "kind": value.kind,
                    "value": normalize_value(value.value, value.numeric_mode),
                    "unit": value.unit,
                    "metric_type": value.metric_type,
                    "is_literal": value.is_literal,
                    "precision": value.precision,
                    "numeric_mode": value.numeric_mode,
                },
                "evidence": normalize_evidence(value.evidence),
            })
        except RetrievalError as error:
            outcomes.append({
                "status": "error",
                "code": error.code,
                "message": error.message,
                "evidence": error.evidence,
            })
        except Exception as error:  # unexpected: surfaced for parity visibility
            outcomes.append({
                "status": "unexpected",
                "code": type(error).__name__,
                "message": str(error),
                "evidence": {},
            })
    return outcomes, sampled


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
    parser.add_argument("--limit", type=int, default=8)
    parser.add_argument("--workspace", type=Path, default=ROOT)
    parser.add_argument("--report", type=Path, required=True)
    args = parser.parse_args()

    source_root = args.source_root.resolve(strict=True)
    workspace = args.workspace.resolve(strict=True)
    report_path = args.report.resolve()
    sys.path.insert(0, str(source_root / "core" / "src"))

    with tempfile.TemporaryDirectory(prefix="audit-s5-retrieve-") as temporary:
        directory = Path(temporary)
        config_copy = directory / "config.xlsx"
        audit_copy = directory / "source.xlsx"
        shutil.copy2(args.config_workbook.resolve(strict=True), config_copy)
        shutil.copy2(args.audit_copy.resolve(strict=True), audit_copy)

        expected_calls, sampled_calls = python_calls(
            directory, config_copy, audit_copy, args.limit
        )

        requests = []
        for index, (rule, args, keywords) in enumerate(sampled_calls):
            requests.append({
                "id": index + 1,
                "source": str(audit_copy),
                "business": rule.business,
                "default_sheet": rule.default_sheet,
                "calls": [{
                    "name": "V_CUR",
                    "args": [
                        list(argument) if isinstance(argument, tuple) else str(argument)
                        for argument in args
                    ],
                    "keywords": keywords,
                }],
            })
        input_path = directory / "requests.jsonl"
        input_path.write_text(
            "".join(json.dumps(request, ensure_ascii=False) + "\n" for request in requests),
            encoding="utf-8",
        )
        rust_report = directory / "rust_outcomes.jsonl"
        environment = os.environ.copy()
        environment["AUDIT_S5_RETRIEVE_INPUT"] = str(input_path)
        environment["AUDIT_S5_RETRIEVE_REPORT"] = str(rust_report)
        completed = subprocess.run(
            [
                "cargo", "test", "-p", "audit-core", "--test", "s5_retrieve_parity",
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
        rust_outcomes = []
        for line in rust_report.read_text(encoding="utf-8").splitlines():
            if not line.strip():
                continue
            payload = json.loads(line)
            rust_outcomes.extend(payload.get("calls", []))

        matched = 0
        mismatch_indexes: list[int] = []
        error_count = 0
        for index, expected in enumerate(expected_calls):
            actual = rust_outcomes[index] if index < len(rust_outcomes) else {"status": "missing"}
            if actual.get("result", {}).get("value") is not None:
                actual["result"]["value"] = actual["result"]["value"].get("value")
            diffs: list[str] = []
            if expected["status"] != actual.get("status"):
                diffs.append("status")
                if len(mismatch_indexes) < 3:
                    print(json.dumps({
                        "dbg_status": {"index": index, "python": expected["status"],
                                        "rust": actual.get("status"),
                                        "rust_code": actual.get("code"),
                                        "rust_message": actual.get("message")},
                    }, ensure_ascii=False))
            elif expected["status"] == "ok":
                diff(expected["result"], actual.get("result", {}), "result", diffs)
                diff(
                    {"evidence": expected["evidence"]},
                    {"evidence": actual.get("evidence", [])},
                    "evidence",
                    diffs,
                )
            else:
                if expected.get("code") != actual.get("code"):
                    diffs.append("code")
                if expected.get("message") != actual.get("message"):
                    diffs.append("message")
            if expected["status"] != "ok":
                error_count += 1
            if diffs:
                mismatch_indexes.append(index)
            else:
                matched += 1

    report = {
        "schema_version": 1,
        "suite": "audit.s5_retrieve_parity",
        "status": "completed" if not mismatch_indexes else "mismatch",
        "sample_scope": {
            "rule_call_samples": len(expected_calls),
            "business": "单位贷款",
            "audit_period": "2026-07-31",
            "files_copied_to_temporary": True,
        },
        "comparison": {
            "matched": matched,
            "mismatch_indexes": mismatch_indexes[:20],
            "mismatch_count": len(mismatch_indexes),
            "error_cases": error_count,
        },
        "verification": {
            "rust_runner": "audit-core/tests/s5_retrieve_parity.rs",
            "python_baseline": "s1audit_configured.sources.SourceRepository.retrieve",
            "inputs_unchanged": True,
            "raw_values_persisted": False,
            "capability_status": "unverified",
            "production_route_eligible": False,
        },
        "limitations": [
            "Single-binding current-source retrieval without 数据源契约; the "
            "contract identity/period branches are not exercised by the real "
            "configuration (no contract sheet).",
            "V_EXT external retrieval and the numeric-region template "
            "projection for bound templates remain separate slices.",
        ],
    }
    report_path.parent.mkdir(parents=True, exist_ok=True)
    report_path.write_text(
        json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )
    print(json.dumps({
        "status": report["status"],
        "cases": len(expected_calls),
        "matched": matched,
        "mismatches": len(mismatch_indexes),
    }))
    return 0 if not mismatch_indexes else 1


if __name__ == "__main__":
    raise SystemExit(main())
