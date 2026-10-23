#!/usr/bin/env python3
"""Compare the Rust S1 history CLI with the read-only Python history baseline."""

from __future__ import annotations

import argparse
import json
import subprocess
import sys
import tempfile
from dataclasses import asdict
from pathlib import Path


def load_baseline(source_root: Path):
    core = source_root / "core"
    sys.path.insert(0, str(core / "src"))
    from base_audit.history import classify_current_issues, merge_history
    from base_audit.models import Issue

    return Issue, classify_current_issues, merge_history


def issue(issue_type, issue_id: str, period: str, *, triggered: bool = True):
    return issue_type(
        issue_id=issue_id,
        period=period,
        batch_id="batch-old",
        audit_time="old-time",
        triggered=triggered,
        status="新增",
        first_seen_period=period,
        previous_seen_period="",
        consecutive_count=1,
        org_code="001",
        org_name="测试机构",
        report_code="R01",
        sheet_name="Sheet1",
        rule_id=issue_id,
        severity="错误",
        formula_cell="C2",
        target_cell="B2",
        target_value=1,
        formula_result="错误",
        message="测试",
        source_file="source.xlsx",
        audit_file="audit.xlsx",
    )


def run_rust(binary: Path, input_path: Path, developer_view: bool):
    command = [str(binary), "s1-history-classify", "--input", str(input_path)]
    if developer_view:
        command.append("--developer-view")
    completed = subprocess.run(command, check=True, capture_output=True, text=True)
    return json.loads(completed.stdout), completed.stderr


def first_difference(expected, actual, path="$",):
    if type(expected) is not type(actual):
        return path, expected, actual
    if isinstance(expected, dict):
        for key in sorted(set(expected) | set(actual)):
            if key not in expected or key not in actual:
                return f"{path}.{key}", expected.get(key), actual.get(key)
            difference = first_difference(expected[key], actual[key], f"{path}.{key}")
            if difference:
                return difference
    elif isinstance(expected, list):
        for index, (left, right) in enumerate(zip(expected, actual)):
            difference = first_difference(left, right, f"{path}[{index}]")
            if difference:
                return difference
        if len(expected) != len(actual):
            return f"{path}.length", len(expected), len(actual)
    elif expected != actual:
        return path, expected, actual
    return None


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--audit-binary", required=True, type=Path)
    parser.add_argument("--source-root", required=True, type=Path)
    parser.add_argument("--report", required=True, type=Path)
    args = parser.parse_args()
    Issue, classify_current_issues, merge_history = load_baseline(args.source_root)

    continuous_first = issue(Issue, "CONT", "2025-12")
    continuous_first.first_seen_period = "2025-12"
    continuous_middle = issue(Issue, "CONT", "2026-01")
    continuous_middle.first_seen_period = "2025-12"
    continuous_last = issue(Issue, "CONT", "2026-02")
    continuous_last.first_seen_period = "2025-12"
    continuous_last.consecutive_count = 2
    continuous_last.institution_feedback = "最近一次机构说明"
    continuous_last.auditor_opinion = "最近一次审核意见"

    recurring = issue(Issue, "AGAIN", "2026-01")
    resolved = issue(Issue, "RESOLVED", "2026-02")
    out_of_scope = issue(Issue, "OUTSIDE", "2026-02")
    out_of_scope.org_code = "999"
    current_period_note_a = issue(Issue, "NOTES", "2026-03")
    current_period_note_a.institution_feedback = " 机构说明甲 "
    current_period_note_b = issue(Issue, "NOTES", "2026-03")
    current_period_note_b.institution_feedback = "机构说明乙"
    current_period_note_b.auditor_opinion = "审核意见"

    history = [
        continuous_first,
        continuous_middle,
        continuous_last,
        recurring,
        resolved,
        out_of_scope,
        current_period_note_a,
        current_period_note_b,
    ]
    current = [
        issue(Issue, "CONT", "2026-03"),
        issue(Issue, "AGAIN", "2026-03"),
        issue(Issue, "NEW", "2026-03"),
        issue(Issue, "NOTES", "2026-03"),
    ]
    request = {
        "period": "2026-03",
        "batch_id": "batch-new",
        "audit_time": "new-time",
        "audited_org_codes": ["001"],
        "current_issues": [asdict(item) for item in current],
        "history": [asdict(item) for item in history],
    }

    merged_history = merge_history(history, current)
    classified, resolved_issues = classify_current_issues(
        "2026-03",
        current,
        history,
        audited_org_codes={"001"},
        batch_id="batch-new",
        audit_time="new-time",
    )
    expected = {
        "current_issues": [asdict(item) for item in classified],
        "resolved_issues": [asdict(item) for item in resolved_issues],
        "merged_history": [asdict(item) for item in merged_history],
    }

    args.report.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="audit-s1-history-") as temporary:
        input_path = Path(temporary) / "history-input.json"
        input_path.write_text(json.dumps(request, ensure_ascii=False), encoding="utf-8")
        ordinary_output, ordinary_log = run_rust(args.audit_binary, input_path, False)
        developer_output, developer_log = run_rust(args.audit_binary, input_path, True)

    compared_fields = ("current_issues", "resolved_issues", "merged_history")
    differences = [
        field
        for field in compared_fields
        if ordinary_output[field] != expected[field] or developer_output[field] != expected[field]
    ]
    if differences:
        field = differences[0]
        detail = first_difference(expected[field], ordinary_output[field], f"$.{field}")
        raise AssertionError(f"Python/Rust history field differences: {differences}; first={detail}")
    if "引擎 Rust" not in ordinary_log or "耗时" not in ordinary_log:
        raise AssertionError("ordinary user log omits engine or elapsed time")
    developer_events = [json.loads(line) for line in developer_log.splitlines() if line.strip()]
    if not any(
        event.get("step") == "history_issue_classification"
        and event.get("status") == "completed"
        and event.get("engine") == "rust_s1_history"
        and "elapsed_ms" in event
        for event in developer_events
    ):
        raise AssertionError("developer log omits completed engine/timing event")

    summary = ordinary_output["summary"]
    report = {
        "suite": "audit.s1_history_classification",
        "baseline": "Open-Data-Audit base_audit.history",
        "status": "passed",
        "fixture": "synthetic_issue_history",
        "history_rows": len(history),
        "current_issues": len(classified),
        "resolved_issues": len(resolved_issues),
        "merged_history_rows": len(expected["merged_history"]),
        "python_rust_field_differences": 0,
        "compared_fields": list(compared_fields),
        "summary": {key: summary[key] for key in (
            "new_issues", "continuous_issues", "recurring_issues", "resolved_issues", "previous_period"
        )},
        "user_log_engine_and_timing": True,
        "developer_log_engine_and_timing": True,
        "source_workbooks_modified": False,
    }
    args.report.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
