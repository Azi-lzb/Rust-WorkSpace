#!/usr/bin/env python3
"""Compare the Rust S1 issue chain with the read-only Python workflow stages."""

from __future__ import annotations

import argparse
import dataclasses
import hashlib
import json
import os
import subprocess
import sys
import tempfile
import time
from collections import Counter
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path
from typing import Any

sys.dont_write_bytecode = True


def wsl_path(path: Path) -> str:
    drive = path.drive.rstrip(":").lower()
    if len(drive) != 1:
        raise ValueError(f"无法转换为 WSL 挂载路径：{path.drive}")
    relative = path.as_posix().split("/", 1)[1]
    return f"/mnt/{drive}/{relative}"


def canonical(value: Any) -> Any:
    if isinstance(value, Decimal):
        return value
    if isinstance(value, bool) or value is None:
        return value
    if isinstance(value, (int, float)):
        return Decimal(str(value))
    if isinstance(value, (date, datetime)):
        return value.isoformat()
    if isinstance(value, dict):
        return {str(key): canonical(item) for key, item in value.items()}
    if isinstance(value, (list, tuple)):
        return [canonical(item) for item in value]
    return value


def normalized_issue(issue: dict[str, Any]) -> dict[str, Any]:
    row = canonical(issue)
    for field in ("source_file", "audit_file"):
        value = row.get(field)
        if value:
            row[field] = Path(str(value)).name.casefold()
    return row


def compare_issues(expected: list[Any], actual: list[dict[str, Any]]) -> dict[str, Any]:
    expected_rows = {
        issue.issue_id: normalized_issue(dataclasses.asdict(issue))
        for issue in expected
    }
    actual_rows = {
        issue["issue_id"]: normalized_issue(issue)
        for issue in actual
    }
    if len(expected_rows) != len(expected) or len(actual_rows) != len(actual):
        raise ValueError("Issue identity is not unique; refusing to hide duplicate records")
    left, right = set(expected_rows), set(actual_rows)
    differences: Counter[str] = Counter()
    for issue_id in left & right:
        expected_item = expected_rows[issue_id]
        actual_item = actual_rows[issue_id]
        for field in sorted(set(expected_item) | set(actual_item)):
            if expected_item.get(field) != actual_item.get(field):
                differences[field] += 1
    return {
        "expected_count": len(expected),
        "actual_count": len(actual),
        "missing_in_rust": len(left - right),
        "missing_in_python": len(right - left),
        "field_difference_counts": dict(sorted(differences.items())),
    }


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def invoke_rust(args: argparse.Namespace, manifest: Path, source: Path) -> tuple[dict[str, Any], str, int]:
    binary = args.audit_binary.resolve(strict=True)
    if os.name == "nt" and binary.suffix.lower() != ".exe":
        command = [
            "wsl.exe",
            "--distribution",
            args.wsl_distribution,
            "--user",
            "root",
            "--exec",
            wsl_path(binary),
        ]
        input_arg = wsl_path(args.input_dir.resolve(strict=True))
        template_arg = wsl_path(args.template_workbook.resolve(strict=True))
        manifest_arg = wsl_path(manifest.resolve(strict=True))
        source_arg = source.relative_to(args.input_dir.resolve(strict=True)).as_posix()
        audit_arg = wsl_path(args.audit_workbook.resolve(strict=True))
        mapping_arg = wsl_path(args.mapping_config.resolve(strict=True)) if args.mapping_config else None
    else:
        command = [str(binary)]
        input_arg = str(args.input_dir.resolve(strict=True))
        template_arg = str(args.template_workbook.resolve(strict=True))
        manifest_arg = str(manifest.resolve(strict=True))
        source_arg = source.relative_to(args.input_dir.resolve(strict=True)).as_posix()
        audit_arg = str(args.audit_workbook.resolve(strict=True))
        mapping_arg = str(args.mapping_config.resolve(strict=True)) if args.mapping_config else None

    command.extend(
        [
            "s1-audit-directory",
            "--input-dir",
            input_arg,
            "--selected-file",
            source_arg,
            "--template-workbook",
            template_arg,
            "--formula-manifest",
            manifest_arg,
            "--period",
            args.period,
            "--batch-id",
            args.batch_id,
            "--audit-time",
            args.audit_time,
            "--developer-view",
        ]
    )
    if mapping_arg:
        command.extend(["--mapping-config", mapping_arg])
    started = time.perf_counter()
    completed = subprocess.run(
        command,
        check=False,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )
    elapsed_ms = round((time.perf_counter() - started) * 1000)
    if completed.returncode:
        raise RuntimeError(
            "Rust S1 issue-chain CLI failed: "
            + completed.stderr.strip().splitlines()[-1]
        )
    output = json.loads(completed.stdout)
    return output, completed.stderr, elapsed_ms


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--source-root", required=True, type=Path)
    parser.add_argument("--audit-binary", required=True, type=Path)
    parser.add_argument("--input-dir", required=True, type=Path)
    parser.add_argument("--source-file", required=True, type=Path)
    parser.add_argument("--audit-workbook", required=True, type=Path)
    parser.add_argument("--template-workbook", required=True, type=Path)
    parser.add_argument("--protected-original", action="append", type=Path, default=[])
    parser.add_argument("--mapping-config", type=Path)
    parser.add_argument("--wsl-distribution", default="Ubuntu")
    parser.add_argument("--period", required=True)
    parser.add_argument("--batch-id", default="s1-chain-parity")
    parser.add_argument("--audit-time", default="2026-10-24 12:00:00")
    parser.add_argument("--report", required=True, type=Path)
    parser.add_argument("--source-filename-alias-used", action="store_true")
    args = parser.parse_args()

    source = args.source_file.resolve(strict=True)
    audit = args.audit_workbook.resolve(strict=True)
    template = args.template_workbook.resolve(strict=True)
    input_dir = args.input_dir.resolve(strict=True)
    source.relative_to(input_dir)
    mapping_config = args.mapping_config.resolve(strict=True) if args.mapping_config else None
    protected_paths = [source, audit, template]
    protected_paths.extend(path.resolve(strict=True) for path in args.protected_original)
    before = {
        path: sha256_file(path) for path in protected_paths
    }

    sys.path.insert(0, str(args.source_root.resolve(strict=True) / "core" / "src"))
    from base_audit.conditional_evaluators import ConditionalFormatService
    from base_audit.name_config import (
        CONDITIONAL_FORMAT_EXTRACT_FUNCTION,
        FORMULA_COPY_FUNCTION,
        ISSUE_EXTRACT_FUNCTION,
        STRUCTURE_COMPARE_FUNCTION,
        features_of_type,
        load_feature_mappings,
    )
    from base_audit.native.history_io import read_history_xlsx
    from base_audit.native.openpyxl_workbook import (
        extract_issues,
        named_ranges,
        read_template,
    )
    from base_audit.service import _organisation_from_name
    from base_audit.history import classify_current_issues, merge_history
    from openpyxl import load_workbook

    mappings = load_feature_mappings(mapping_config, template)
    definition = read_template(
        template,
        formula_mappings=features_of_type(mappings, FORMULA_COPY_FUNCTION),
        structure_mappings=features_of_type(mappings, STRUCTURE_COMPARE_FUNCTION),
        extraction_mappings=features_of_type(mappings, ISSUE_EXTRACT_FUNCTION),
    )
    conditional_mappings = features_of_type(mappings, CONDITIONAL_FORMAT_EXTRACT_FUNCTION)
    template_book = load_workbook(template, read_only=True, data_only=False, keep_links=False)
    try:
        conditional_ranges = named_ranges(template_book, conditional_mappings)
        structure_ranges = definition.structure_ranges
    finally:
        template_book.close()

    org_code, org_name = _organisation_from_name(source)
    formula_issues = extract_issues(
        audit,
        definition.rules,
        period=args.period,
        batch_id=args.batch_id,
        org_code=org_code,
        org_name=org_name,
        source_file=source,
        extraction_ranges=definition.extraction_ranges,
    )
    for issue in formula_issues:
        issue.audit_time = args.audit_time

    conditional_issues, conditional_stats = ConditionalFormatService(mode="PYTHON").extract_issues(
        workbook_path=source,
        ranges=conditional_ranges,
        structure_ranges=structure_ranges,
        period=args.period,
        batch_id=args.batch_id,
        audit_time=args.audit_time,
        source_file=source,
    )
    formula_and_conditional = formula_issues + [
        dataclasses.replace(issue, audit_file=str(audit))
        for issue in conditional_issues
    ]
    python_merged_history = merge_history([], formula_and_conditional)
    python_current, python_resolved = classify_current_issues(
        args.period,
        formula_and_conditional,
        [],
        batch_id=args.batch_id,
        audit_time=args.audit_time,
    )

    audit_manifest_path = (
        wsl_path(audit)
        if os.name == "nt"
        and args.audit_binary.suffix.lower() != ".exe"
        else str(audit)
    )
    with tempfile.TemporaryDirectory(prefix="audit-s1-issue-chain-") as directory:
        manifest = Path(directory) / "formula-copies.json"
        manifest.write_text(
            json.dumps(
                {
                    "files": [
                        {
                            "source_file": source.relative_to(input_dir).as_posix(),
                            "audit_workbook": audit_manifest_path,
                        }
                    ]
                },
                ensure_ascii=False,
            ),
            encoding="utf-8",
        )
        rust, rust_log, rust_elapsed_ms = invoke_rust(args, manifest, source)

    formula_comparison = compare_issues(python_current, rust.get("issues", []))
    resolved_comparison = compare_issues(python_resolved, rust.get("resolved_issues", []))
    history_comparison = compare_issues(python_merged_history, rust.get("merged_history", []))
    actual_formula_count = rust["formula_summary"]["issues"]
    actual_conditional_count = rust["conditional_summary"]["issues"]
    actual_preflight = rust.get("structure_preflight")
    if actual_preflight and actual_preflight["files_accepted"] != 1:
        raise ValueError("Rust structure preflight did not accept the selected real source")
    if actual_formula_count != len(formula_issues):
        raise ValueError("Rust formula issue count differs from the Python formula baseline")
    if actual_conditional_count != len(conditional_issues):
        raise ValueError("Rust conditional issue count differs from the Python baseline")

    after = {
        path: sha256_file(path) for path in protected_paths
    }
    inputs_unchanged = before == after
    all_comparisons = [formula_comparison, resolved_comparison, history_comparison]
    mismatch_count = sum(
        item["missing_in_rust"]
        + item["missing_in_python"]
        + sum(item["field_difference_counts"].values())
        for item in all_comparisons
    )
    log_events = [
        json.loads(line)
        for line in rust_log.splitlines()
        if line.lstrip().startswith("{")
    ]
    actual_readers = sorted(
        {
            event.get("engine") or event.get("result_reader")
            for event in log_events
            if event.get("step")
            in {
                "formula_result_issue_extraction",
                "conditional_format_issue_extraction",
                "issue_merge_and_history_classification",
            }
            and (event.get("status") == "completed" or event.get("engine"))
        }
    )
    report = {
        "schema_version": 1,
        "suite": "audit.s1_issue_chain",
        "baseline": "Open-Data-Audit UOS/native DAG adapter formula/conditional Issue stages and history classifier",
        "status": "passed" if mismatch_count == 0 and inputs_unchanged else "failed",
        "python": {
            "formula_issues": len(formula_issues),
            "conditional_issues": len(conditional_issues),
            "conditional_evaluated_cells": conditional_stats.get("evaluated_cells", 0),
            "merged_issues": len(formula_and_conditional),
        },
        "rust": {
            "formula_issues": actual_formula_count,
            "conditional_issues": actual_conditional_count,
            "merged_issues": len(rust.get("issues", [])),
            "resolved_issues": len(rust.get("resolved_issues", [])),
            "readers": actual_readers,
            "cli_elapsed_ms": rust_elapsed_ms,
        },
        "comparison": {
            "current_issues": formula_comparison,
            "resolved_issues": resolved_comparison,
            "merged_history": history_comparison,
            "mismatch_count": mismatch_count,
        },
        "verification": {
            "all_workbooks_unchanged": inputs_unchanged,
            "formula_recalculation": "not performed; saved audit-copy values only",
            "raw_issue_values_or_paths_persisted": False,
            "python_conditional_metadata_source": "workflow.adapters.native_audit_handlers.handle_conditional_issue_extract",
            "windows_general_handler_metadata_behavior": "workflow.adapters.audit_handlers.handle_conditional_issue_extract passes source organization fields; S1 UOS/native path intentionally leaves them empty",
            "scope": "one read-only real source/audit-copy pair and one S1 template",
            "source_filename_alias_used_for_linux_component_limit": args.source_filename_alias_used,
            "python_conditional_total_time_s": round(
                float(conditional_stats.get("total_time", 0.0)), 3
            ),
        },
    }
    args.report.parent.mkdir(parents=True, exist_ok=True)
    args.report.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0 if report["status"] == "passed" else 1


if __name__ == "__main__":
    raise SystemExit(main())
