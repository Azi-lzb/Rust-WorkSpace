#!/usr/bin/env python3
"""Compare Rust conditional-audit issue rows with the read-only Python service."""

from __future__ import annotations

import argparse
import dataclasses
import datetime as dt
import json
import math
import sys
from collections import Counter
from decimal import Decimal
from pathlib import Path


SOURCE_CORE = Path(r"C:\Users\AZI\Desktop\VibeCode\3.open-data-audit\core")
sys.path.insert(0, str(SOURCE_CORE / "src"))

from base_audit.conditional_evaluators import ConditionalFormatService  # noqa: E402


def canonical(value):
    if dataclasses.is_dataclass(value):
        value = dataclasses.asdict(value)
    if isinstance(value, Decimal):
        return value
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return Decimal(str(value))
    if isinstance(value, (dt.date, dt.datetime, dt.time)):
        return value.isoformat()
    if isinstance(value, float) and not math.isfinite(value):
        return str(value)
    if isinstance(value, dict):
        return {str(key): canonical(item) for key, item in value.items()}
    if isinstance(value, (list, tuple)):
        return [canonical(item) for item in value]
    return value


def normalize_issue(issue):
    row = canonical(issue)
    if not isinstance(row, dict):
        row = vars(issue)
    for key in ("source_file", "audit_file"):
        if row.get(key):
            row[key] = Path(str(row[key])).name
    return row


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--workbook", required=True, type=Path)
    parser.add_argument("--rust-report", required=True, type=Path)
    parser.add_argument("--source-file", type=Path)
    parser.add_argument("--period", required=True)
    parser.add_argument("--batch-id", required=True)
    parser.add_argument("--audit-time", required=True)
    parser.add_argument("--org-code", default="")
    parser.add_argument("--org-name", default="")
    parser.add_argument("--output", type=Path)
    args = parser.parse_args()

    workbook = args.workbook.resolve(strict=True)
    source_file = (args.source_file or workbook).resolve(strict=True)
    service = ConditionalFormatService(mode="PYTHON")
    python_issues, python_stats = service.extract_issues(
        workbook_path=workbook,
        ranges=[],
        structure_ranges=[],
        period=args.period,
        batch_id=args.batch_id,
        audit_time=args.audit_time,
        source_file=source_file,
        org_code=args.org_code,
        org_name=args.org_name,
    )
    rust_report = json.loads(args.rust_report.read_text(encoding="utf-8"))
    python_rows = [normalize_issue(issue) for issue in python_issues]
    rust_rows = [normalize_issue(issue) for issue in rust_report.get("issues", [])]
    python_by_id = {row["issue_id"]: row for row in python_rows}
    rust_by_id = {row["issue_id"]: row for row in rust_rows}
    field_differences: Counter[str] = Counter()
    missing_in_rust = 0
    missing_in_python = 0
    for issue_id in python_by_id.keys() | rust_by_id.keys():
        left = python_by_id.get(issue_id)
        right = rust_by_id.get(issue_id)
        if left is None:
            missing_in_python += 1
            continue
        if right is None:
            missing_in_rust += 1
            continue
        for field in left.keys() | right.keys():
            if left.get(field) != right.get(field):
                field_differences[field] += 1

    mismatch_count = missing_in_rust + missing_in_python + sum(field_differences.values())
    rust_evaluation = rust_report.get("evaluation", {})
    summary = {
        "schema_version": 1,
        "suite": "audit.s1_conditional_issue_extraction",
        "baseline": "Open-Data-Audit ConditionalFormatService(PYTHON)",
        "status": "passed" if mismatch_count == 0 else "failed",
        "python_issue_count": len(python_rows),
        "rust_issue_count": len(rust_rows),
        "missing_in_rust": missing_in_rust,
        "missing_in_python": missing_in_python,
        "field_difference_counts": dict(sorted(field_differences.items())),
        "mismatch_count": mismatch_count,
        "python_evaluated_cells": python_stats.get("evaluated_cells", 0),
        "rust_evaluated_cells": rust_evaluation.get("scanned_cells", 0),
        "python_rules_total": python_stats.get("rules_total", 0),
        "rust_rules_total": rust_evaluation.get("rules_total", 0),
        "python_unsupported_cells": python_stats.get("unsupported_cells", 0),
        "rust_diagnostics": len(rust_evaluation.get("diagnostics", [])),
        "execution": {
            "python_engine": "PythonConditionalRuleEvaluator + Direct OOXML reader",
            "python_rule_read_time_s": python_stats.get("rule_read_time", 0.0),
            "python_rule_normalize_time_s": python_stats.get("rule_normalize_time", 0.0),
            "python_evaluation_time_s": python_stats.get("evaluation_time", 0.0),
            "python_issue_build_time_s": python_stats.get("result_build_time", 0.0),
            "python_total_time_s": python_stats.get("total_time", 0.0),
            "rust_engine": rust_evaluation.get("reader", ""),
            "rust_read_time_ms": rust_evaluation.get("read_time_ms", 0),
            "rust_evaluation_time_ms": rust_evaluation.get("evaluation_time_ms", 0),
            "rust_issue_build_time_ms": rust_report.get("issue_build_time_ms", 0),
            "rust_total_time_ms": rust_evaluation.get("total_time_ms", 0)
            + rust_report.get("issue_build_time_ms", 0),
            "scope": "single real-workbook run; not a performance benchmark",
        },
        "verification": {
            "source_workbook": "read-only",
            "formula_recalculation": "not performed",
            "issue_order": "ignored; issue records are matched by issue_id",
            "source_paths": "compared by basename; not included in output",
        },
    }
    encoded = json.dumps(summary, ensure_ascii=False, indent=2) + "\n"
    if args.output:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(encoded, encoding="utf-8")
    print(encoded, end="")
    return 0 if mismatch_count == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())
