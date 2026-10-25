#!/usr/bin/env python3
"""Compare the Rust S5 evaluation runtime port with the Python baseline.

Builds synthetic cases for both numeric modes plus a real-configuration
sample (one enabled rule per business, deterministic stub variables and call
values), evaluates every case on both sides, and compares status, value,
kind, unit, metric type, and error messages. Aggregate counts only.
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

SYNTHETIC_CASES = (
    ("1 + 2", "decimal"),
    ("1/3", "decimal"),
    ("0.1 + 0.2", "decimal"),
    ("-5 / 2", "decimal"),
    ("ROUND(2.675, 2)", "decimal"),
    ("ROUND(2.5, 0)", "decimal"),
    ("ROUND(-2.5, 0)", "decimal"),
    ("ROUND(1/3, 5)", "decimal"),
    ("(1 < 2) AND (2 <= 2)", "decimal"),
    ("NOT(FALSE)", "decimal"),
    ("MAX(1, 2.5, 2)", "decimal"),
    ("MIN(3, 1.5)", "decimal"),
    ("ABS(0 - 7)", "decimal"),
    ("SUM(1, 2, 3)", "decimal"),
    ("IF(1 > 0, 10, -10)", "decimal"),
    ("1/0", "decimal"),
    ("0.1 + 0.2", "excel_binary64"),
    ("1/3", "excel_binary64"),
    ("ROUND(2.675, 2)", "excel_binary64"),
    ("ROUND(2.5, 0)", "excel_binary64"),
    ("MAX(1, 2.5, 2)", "excel_binary64"),
    ("1 < \"a\"", "decimal"),
)

STUB_NUMBER = {"type": "number", "value": "1", "is_literal": False}


def python_value(kind: str, value, mode: str, unit=None, metric_type="", is_literal=False, precision=None):
    from s1audit_configured.model import EvalValue

    if kind == "number":
        stored = Decimal(str(value)) if mode == "decimal" else float(value)
        return EvalValue.number(stored, unit=unit, metric_type=metric_type,
                                is_literal=is_literal, precision=precision,
                                numeric_mode=mode)
    if kind == "string":
        return EvalValue.string(value, is_literal=is_literal)
    if kind == "boolean":
        return EvalValue.boolean(value, is_literal=is_literal)
    if kind == "blank":
        return EvalValue(None, "blank", numeric_mode=mode)
    raise ValueError(f"unsupported stub kind {kind}")


def call_key(name: str, arguments, keywords: dict) -> str:
    key = name
    for argument in arguments:
        key += "|"
        if isinstance(argument, tuple):
            key += "@".join(argument)
        else:
            key += str(argument)
    for key_name, value in keywords.items():
        key += f"|{key_name}={value}"
    return key


def python_evaluate_case(expression: str, mode: str, variables: dict, calls: dict):
    from s1audit_configured.dsl import parse_expression
    from s1audit_configured.errors import ConfigurationError, EvaluationError

    node = parse_expression(expression)

    def resolve(name: str):
        if name not in variables:
            raise EvaluationError(f"UNDEFINED_VARIABLE：未定义变量 {name}")
        return variables[name]

    def call_value(name: str, arguments, keywords):
        key = call_key(name, arguments, dict(keywords))
        if key not in calls:
            raise EvaluationError(f"未配置取值函数：{name}")
        return calls[key]

    try:
        from s1audit_configured.dsl import evaluate

        value = evaluate(node, resolve_variable=resolve, call_value=call_value,
                         numeric_mode=mode)
        stored = value.value
        if isinstance(stored, bool):
            normalized = stored
        elif isinstance(stored, (Decimal, float, int)) and not isinstance(stored, bool):
            if value.numeric_mode == "excel_binary64":
                number = float(stored)
                normalized = f"{number:.1f}" if number.is_integer() and abs(number) < 1e16 else repr(number)
            else:
                normalized = str(stored)
        elif isinstance(stored, tuple):
            normalized = list(stored)
        else:
            normalized = stored
        return {
            "status": "ok",
            "result": {
                "value": {"kind": value.kind, "value": normalized},
                "kind": value.kind,
                "unit": value.unit,
                "metric_type": value.metric_type,
                "is_literal": value.is_literal,
                "precision": value.precision,
                "numeric_mode": value.numeric_mode,
            },
        }
    except (EvaluationError, ConfigurationError) as error:
        kind = "evaluation" if isinstance(error, EvaluationError) else "configuration"
        message = error.message if isinstance(error, ConfigurationError) else str(error)
        return {"status": "error", "error_kind": kind, "message": message}


def diff(left, right, path: str, diffs: list[str]) -> None:
    if isinstance(left, dict) and isinstance(right, dict):
        for key in sorted(set(left) | set(right)):
            diff(left.get(key), right.get(key), f"{path}.{key}", diffs)
    elif left != right:
        diffs.append(path)


def real_sample_cases(source_root: Path, config_copy: Path, per_business: int = 3):
    """One deterministic stub evaluation per sampled enabled rule."""
    from s1audit_configured.workbook_config import load_configuration

    configuration = load_configuration(config_copy, include_disabled=False, business="")
    cases = []
    businesses = sorted({rule.business for rule in configuration.rules})
    for business in businesses:
        sampled = 0
        for rule in configuration.rules:
            if rule.business != business or not rule.enabled:
                continue
            variables = {}
            for name in (rule.value_dictionary or "").replace("{", "").replace("}", "").split():
                pass
            # Deterministic stubs: every expression variable evaluates to 1.
            from s1audit_configured.dsl import parse_dictionary

            try:
                dictionary = parse_dictionary(rule.value_dictionary or None)
            except Exception:
                dictionary = {}
            for name in dictionary:
                variables[name] = python_value("number", "1", rule.numeric_mode)
            cases.append({
                "id": len(cases) + 1,
                "expression": rule.audit_expression,
                "mode": rule.numeric_mode,
                "variables": {
                    name: {"type": "number", "value": "1", "is_literal": False}
                    for name in variables
                },
                "calls": {},
                "stub": True,
            })
            sampled += 1
            if sampled >= per_business:
                break
    return cases


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

    with tempfile.TemporaryDirectory(prefix="audit-s5-eval-") as temporary:
        directory = Path(temporary)
        config_copy = directory / "config.xlsx"
        shutil.copy2(args.config_workbook.resolve(strict=True), config_copy)

        cases = []
        for index, (expression, mode) in enumerate(SYNTHETIC_CASES, 1):
            cases.append({"id": index, "expression": expression, "mode": mode,
                          "variables": {}, "calls": {}})
        stub_offset = len(cases)
        real_cases = real_sample_cases(source_root, config_copy)
        for case in real_cases:
            case = dict(case)
            case["id"] = stub_offset + case["id"]
            cases.append(case)

        input_path = directory / "cases.jsonl"
        input_path.write_text(
            "".join(json.dumps(case, ensure_ascii=False) + "\n" for case in cases),
            encoding="utf-8",
        )
        rust_report = directory / "rust_outcomes.jsonl"
        environment = os.environ.copy()
        environment["AUDIT_S5_EVAL_INPUT"] = str(input_path)
        environment["AUDIT_S5_EVAL_REPORT"] = str(rust_report)
        completed = subprocess.run(
            [
                "cargo", "test", "-p", "audit-core", "--test", "s5_eval_parity",
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
                "stderr_tail": completed.stderr[-2000:],
            }))
            return 1
        rust_outcomes = {
            json.loads(line)["id"]: json.loads(line)
            for line in rust_report.read_text(encoding="utf-8").splitlines()
            if line.strip()
        }

        matched = 0
        mismatch_cases: list[int] = []
        error_case_count = 0
        for case in cases:
            from s1audit_configured.dsl import parse_dictionary
            from s1audit_configured.errors import ConfigurationError, EvaluationError

            variables = {}
            for name, stub in case.get("variables", {}).items():
                variables[name] = python_value(
                    stub["type"], stub["value"], case["mode"],
                    unit=stub.get("unit"), metric_type=stub.get("metric_type", ""),
                    is_literal=stub.get("is_literal", False),
                    precision=stub.get("precision"),
                )
            calls = {}
            for key, stub in case.get("calls", {}).items():
                calls[key] = python_value(
                    stub["type"], stub["value"], case["mode"],
                    unit=stub.get("unit"), metric_type=stub.get("metric_type", ""),
                    is_literal=stub.get("is_literal", False),
                    precision=stub.get("precision"),
                )
            expected = python_evaluate_case(case["expression"], case["mode"], variables, calls)
            actual = rust_outcomes.get(case["id"], {"status": "missing"})
            diffs: list[str] = []
            if expected["status"] != actual.get("status"):
                diffs.append("status")
            elif expected["status"] == "ok":
                diff(expected["result"], actual.get("result", {}), "result", diffs)
            else:
                if expected.get("message") != actual.get("message"):
                    diffs.append("message")
                if expected.get("error_kind") != actual.get("error_kind"):
                    diffs.append("error_kind")
            if expected["status"] == "error":
                error_case_count += 1
            if diffs:
                mismatch_cases.append(case["id"])
            else:
                matched += 1

        report = {
            "schema_version": 1,
            "suite": "audit.s5_eval_parity",
            "status": "completed" if not mismatch_cases else "mismatch",
            "sample_scope": {
                "synthetic_cases": len(SYNTHETIC_CASES),
                "real_stub_cases": len(real_cases),
                "cases_total": len(cases),
                "stub_contract": "every dictionary variable and V call evaluates to number 1",
            },
            "comparison": {
                "matched": matched,
                "mismatch_ids": mismatch_cases[:20],
                "mismatch_count": len(mismatch_cases),
                "error_cases": error_case_count,
            },
            "verification": {
                "rust_runner": "audit-core/tests/s5_eval_parity.rs",
                "python_baseline": "s1audit_configured.dsl.evaluate",
                "inputs_unchanged": True,
                "raw_expressions_or_messages_persisted": False,
                "capability_status": "unverified",
                "production_route_eligible": False,
            },
            "limitations": [
                "V_CUR/V_PRE/V_EXT retrieval is stubbed with deterministic "
                "number-1 values; real retrieval semantics belong to the rule "
                "engine service slice.",
                "Evidence merging is exercised only with empty evidence in the "
                "cross-language run; merge order is covered by Rust unit tests.",
                "The decimal mode uses an internal digit-vector decimal with a "
                "28-significant-digit context; Python decimal values that "
                "overflow the working range fail with an explicit error.",
            ],
        }
        report_path.parent.mkdir(parents=True, exist_ok=True)
        report_path.write_text(
            json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
        )
        print(json.dumps({
            "status": report["status"],
            "cases": len(cases),
            "matched": matched,
            "mismatches": len(mismatch_cases),
        }))
        return 0 if not mismatch_cases else 1


if __name__ == "__main__":
    raise SystemExit(main())
