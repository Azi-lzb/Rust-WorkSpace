#!/usr/bin/env python3
"""Compare one Rust S1 structure-preflight result with the read-only Python baseline."""

from __future__ import annotations

import argparse
import json
import math
import sys
import time
from pathlib import Path


sys.dont_write_bytecode = True


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--source-root", type=Path, required=True)
    parser.add_argument("--source-workbook", type=Path, required=True)
    parser.add_argument("--template-workbook", type=Path, required=True)
    parser.add_argument("--rust-report", type=Path, required=True)
    parser.add_argument("--conditional-parity", type=Path)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()

    sys.path.insert(0, str(args.source_root / "core" / "src"))
    from base_audit.models import TemplateDefinition
    from base_audit.name_config import (
        CONDITIONAL_FORMAT_EXTRACT_FUNCTION,
        STRUCTURE_COMPARE_FUNCTION,
        features_of_type,
        load_feature_mappings,
    )
    from base_audit.native.openpyxl_workbook import named_ranges
    from base_audit.native.openpyxl_workbook import template_structure_values
    from base_audit.preflight_xlsx import validate_source_xlsx
    from openpyxl import load_workbook

    source_path = args.source_workbook.resolve(strict=True)
    template_path = args.template_workbook.resolve(strict=True)
    rust_payload = json.loads(args.rust_report.read_text(encoding="utf-8"))
    rust_preflight = rust_payload.get("structure_preflight")
    if not isinstance(rust_preflight, dict):
        raise ValueError("Rust report has no structure_preflight summary")
    rust_files = rust_preflight.get("files", [])
    if len(rust_files) != 1:
        raise ValueError("Rust structure_preflight must contain exactly one selected source")
    rust_match = rust_files[0]

    mappings = load_feature_mappings(None, template_path)
    conditional_mappings = features_of_type(mappings, CONDITIONAL_FORMAT_EXTRACT_FUNCTION)
    structure_mappings = features_of_type(mappings, STRUCTURE_COMPARE_FUNCTION)
    if not conditional_mappings or not structure_mappings:
        raise ValueError("Python default mappings do not contain S1-F01 structure/scan areas")

    template_book = load_workbook(
        template_path, read_only=True, data_only=False, keep_links=False
    )
    try:
        scan_ranges = named_ranges(template_book, conditional_mappings)
        structure_ranges = named_ranges(template_book, structure_mappings)
        definition = TemplateDefinition(
            rules=[],
            copy_ranges=[*scan_ranges, *structure_ranges],
            structured=False,
            structure_ranges=structure_ranges,
        )
        expected = template_structure_values(
            template_path, definition, workbook=template_book
        )
        fixed_label_count = sum(
            1
            for _, rows in expected
            for row in rows
            for value in row
            if value is not None and not (isinstance(value, str) and not value.strip())
        )
    finally:
        template_book.close()

    python_started = time.perf_counter()
    python_match = validate_source_xlsx(source_path, definition, expected)
    python_elapsed_ms = round((time.perf_counter() - python_started) * 1000)

    python_missing = sorted(python_match.missing_sheets)
    rust_missing = sorted(rust_match.get("missing_sheets", []))
    comparisons = {
        "matched": python_match.matched == (rust_match.get("status") == "matched"),
        "matched_labels": python_match.matched_labels == rust_match.get("matched_labels"),
        "checked_labels": python_match.checked_labels == rust_match.get("checked_labels"),
        "score": math.isclose(
            python_match.score,
            float(rust_match.get("score", math.nan)),
            rel_tol=0.0,
            abs_tol=1e-12,
        ),
        "missing_sheets": python_missing == rust_missing,
        "template_label_count": fixed_label_count == rust_preflight.get("template_labels"),
    }
    mismatch_count = sum(not equal for equal in comparisons.values())
    report = {
        "schema_version": 1,
        "suite": "audit.s1_structure_preflight",
        "baseline": "Open-Data-Audit base_audit.preflight_xlsx.validate_source_xlsx",
        "status": "passed" if mismatch_count == 0 else "mismatch",
        "sample": {
            "source_workbook": "redacted real S1 workbook copy",
            "template_workbook": "redacted real S1 template copy",
            "read_only": True,
        },
        "template": {
            "structure_areas": len(structure_ranges),
            "fixed_labels": fixed_label_count,
        },
        "python": {
            "matched": python_match.matched,
            "matched_labels": python_match.matched_labels,
            "checked_labels": python_match.checked_labels,
            "score": python_match.score,
            "missing_sheet_count": len(python_missing),
            "elapsed_ms": python_elapsed_ms,
        },
        "rust": {
            "matched": rust_match.get("status") == "matched",
            "matched_labels": rust_match.get("matched_labels"),
            "checked_labels": rust_match.get("checked_labels"),
            "score": rust_match.get("score"),
            "missing_sheet_count": len(rust_missing),
            "engine": rust_preflight.get("engine"),
            "read_time_ms": rust_match.get("read_time_ms"),
            "elapsed_ms": rust_match.get("elapsed_ms"),
        },
        "comparison": {
            "fields": comparisons,
            "mismatch_count": mismatch_count,
        },
        "verification": {
            "source_and_template_modified": False,
            "cell_values_and_paths_recorded": False,
            "formula_recalculation": "not performed",
            "scope": "one real S1 source/template pair; structure preflight and conditional issue extraction",
        },
    }
    if args.conditional_parity:
        issue_parity = json.loads(args.conditional_parity.read_text(encoding="utf-8"))
        report["conditional_audit"] = {
            "baseline": issue_parity.get("baseline"),
            "status": issue_parity.get("status"),
            "python_issue_count": issue_parity.get("python_issue_count"),
            "rust_issue_count": issue_parity.get("rust_issue_count"),
            "mismatch_count": issue_parity.get("mismatch_count"),
            "field_difference_counts": issue_parity.get("field_difference_counts"),
            "python_evaluated_cells": issue_parity.get("python_evaluated_cells"),
            "rust_evaluated_cells": issue_parity.get("rust_evaluated_cells"),
            "python_rules_total": issue_parity.get("python_rules_total"),
            "rust_rules_total": issue_parity.get("rust_rules_total"),
            "rust_engine": issue_parity.get("execution", {}).get("rust_engine"),
        }
        if issue_parity.get("status") != "passed":
            report["status"] = "mismatch"
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(
        json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0 if report["status"] == "passed" else 1


if __name__ == "__main__":
    raise SystemExit(main())
