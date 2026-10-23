#!/usr/bin/env python3
"""Compare Rust conditional-audit issue rows with the read-only Python service."""

from __future__ import annotations

import argparse
import dataclasses
import datetime as dt
import json
import math
import sys
import tempfile
from collections import Counter
from decimal import Decimal
from pathlib import Path


SOURCE_CORE = Path(r"C:\Users\AZI\Desktop\VibeCode\3.open-data-audit\core")
sys.dont_write_bytecode = True
sys.path.insert(0, str(SOURCE_CORE / "src"))

from base_audit.conditional_evaluators import ConditionalFormatService  # noqa: E402
from base_audit.discovery import recommend_template  # noqa: E402
from base_audit.name_config import (  # noqa: E402
    CONDITIONAL_FORMAT_EXTRACT_FUNCTION,
    STRUCTURE_COMPARE_FUNCTION,
    features_of_type,
    load_feature_mappings,
)
from base_audit.native.openpyxl_workbook import named_ranges  # noqa: E402
from openpyxl import load_workbook  # noqa: E402


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


def count_records(value):
    if isinstance(value, (list, dict)):
        return len(value)
    return int(value or 0)


def template_area_plan(template_path: Path | None, config_path: Path | None):
    if template_path is None:
        return None, [], []
    mappings = load_feature_mappings(config_path, template_path)
    conditional = next(
        (mapping for mapping in mappings
         if mapping.feature_type == CONDITIONAL_FORMAT_EXTRACT_FUNCTION),
        None,
    )
    if conditional is None:
        raise ValueError("模板配置中没有条件格式提取映射")
    structure_mappings = features_of_type(mappings, STRUCTURE_COMPARE_FUNCTION)
    template = load_workbook(
        template_path, read_only=True, data_only=False, keep_links=False,
    )
    try:
        ranges = named_ranges(template, (conditional,))
        structure_ranges = [
            area
            for mapping in structure_mappings
            for area in named_ranges(template, (mapping,))
        ]
    finally:
        template.close()
    return conditional, ranges, structure_ranges


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--workbook", required=True, type=Path)
    parser.add_argument("--rust-report", required=True, type=Path)
    parser.add_argument("--source-file", type=Path)
    parser.add_argument("--template-workbook", type=Path)
    parser.add_argument("--template-dir", type=Path)
    parser.add_argument("--mapping-config", type=Path)
    parser.add_argument("--rust-progress", type=Path)
    parser.add_argument("--period", required=True)
    parser.add_argument("--batch-id", required=True)
    parser.add_argument("--audit-time", required=True)
    parser.add_argument("--org-code", default="")
    parser.add_argument("--org-name", default="")
    parser.add_argument("--output", type=Path)
    args = parser.parse_args()

    workbook = args.workbook.resolve(strict=True)
    source_file = (args.source_file or workbook).resolve(strict=True)
    mapping, ranges, structure_ranges = template_area_plan(
        args.template_workbook.resolve(strict=True) if args.template_workbook else None,
        args.mapping_config.resolve(strict=True) if args.mapping_config else None,
    )
    python_template_recommendation = None
    template_recommendation_matches_python = None
    if args.template_dir:
        with tempfile.TemporaryDirectory(prefix="audit-template-recommendation-compare-") as temp:
            suggestion = recommend_template(
                args.template_dir.resolve(strict=True),
                workbook.parent,
                Path(temp) / "template-index.json",
            )
        python_template_recommendation = {
            "template": (
                suggestion.template_path.name if suggestion.template_path is not None else None
            ),
            "confidence": suggestion.confidence,
            "matched": suggestion.matched,
            "details": suggestion.details,
            "alternatives": [list(item) for item in suggestion.alternatives],
        }
    service = ConditionalFormatService(mode="PYTHON")
    python_issues, python_stats = service.extract_issues(
        workbook_path=workbook,
        ranges=ranges,
        structure_ranges=structure_ranges,
        period=args.period,
        batch_id=args.batch_id,
        audit_time=args.audit_time,
        source_file=source_file,
        mapping=mapping,
        org_code=args.org_code,
        org_name=args.org_name,
    )
    rust_report = json.loads(args.rust_report.read_text(encoding="utf-8"))
    rust_area_mapping = None
    rust_source_discovery = None
    rust_template_recommendation = None
    template_selection_matches_expected = None
    if args.rust_progress:
        progress_events = [
            json.loads(line)
            for line in args.rust_progress.read_text(encoding="utf-8").splitlines()
            if line.strip()
        ]
        mapping_event = next(
            (
                event
                for event in progress_events
                if event.get("step") == "conditional_area_mapping"
                and event.get("status") == "completed"
            ),
            None,
        )
        if args.template_workbook and mapping_event is None:
            raise ValueError("Rust progress has no completed template area mapping event")
        if mapping_event is not None:
            rust_area_mapping = {
                "engine": mapping_event.get("engine", ""),
                "area_source": mapping_event.get("area_source", ""),
                "mappings_loaded": mapping_event.get("mappings_loaded", 0),
                "defined_names_total": mapping_event.get("defined_names_total", 0),
                "mapped_named_areas": mapping_event.get("mapped_named_areas", 0),
                "unresolved_matching_names": mapping_event.get(
                    "unresolved_matching_names", 0
                ),
                "scan_areas": mapping_event.get("scan_areas", 0),
                "structure_areas": mapping_event.get("structure_areas", 0),
                "elapsed_ms": mapping_event.get("elapsed_ms", 0),
            }
        discovery_event = next(
            (
                event
                for event in progress_events
                if event.get("step") == "source_workbook_discovery"
                and event.get("status") == "completed"
            ),
            None,
        )
        if discovery_event is not None:
            rust_source_discovery = {
                "available_source_files": discovery_event.get("available_source_files", 0),
                "selected_files_requested": discovery_event.get("selected_files_requested", 0),
                "selection_applied": discovery_event.get("selection_applied", False),
                "files_total": discovery_event.get("files_total", 0),
                "period_provided": bool(discovery_event.get("period", "")),
                "period_source": discovery_event.get("period_source", ""),
                "elapsed_ms": discovery_event.get("elapsed_ms", 0),
            }
        template_event = next(
            (
                event
                for event in progress_events
                if event.get("step") == "template_recommendation"
                and event.get("status") == "completed"
            ),
            None,
        )
        if template_event is not None:
            selected_template = template_event.get("selected_template", "")
            template_selection_matches_expected = (
                selected_template == args.template_workbook.name
                if args.template_workbook
                else None
            )
            rust_template_recommendation = {
                "engine": template_event.get("engine", ""),
                "source_workbooks": template_event.get("source_workbooks", 0),
                "matched": template_event.get("matched", False),
                "confidence": template_event.get("confidence", 0.0),
                "selected_template": selected_template,
                "details": template_event.get("details", ""),
                "alternatives": template_event.get("alternatives", []),
                "elapsed_ms": template_event.get("elapsed_ms", 0),
                "matches_expected_template": template_selection_matches_expected,
            }
            if python_template_recommendation is not None:
                rust_recommendation_result = {
                    "template": selected_template,
                    "confidence": rust_template_recommendation["confidence"],
                    "matched": rust_template_recommendation["matched"],
                    "details": rust_template_recommendation["details"],
                    "alternatives": rust_template_recommendation["alternatives"],
                }
                template_recommendation_matches_python = (
                    rust_recommendation_result == python_template_recommendation
                )
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
    rust_directory_files = rust_report.get("files", [])
    if not rust_evaluation and rust_directory_files:
        rust_evaluation = {
            "reader": rust_directory_files[0].get("reader", ""),
            "read_time_ms": sum(item.get("read_time_ms", 0) for item in rust_directory_files),
            "evaluation_time_ms": sum(
                item.get("evaluation_time_ms", 0) for item in rust_directory_files
            ),
            "total_time_ms": sum(item.get("total_time_ms", 0) for item in rust_directory_files),
            "scanned_cells": sum(item.get("scanned_cells", 0) for item in rust_directory_files),
            "rules_total": sum(item.get("rules_total", 0) for item in rust_directory_files),
            "rules_supported": sum(
                item.get("rules_supported", 0) for item in rust_directory_files
            ),
            "_diagnostics_count": sum(
                count_records(item.get("diagnostics", 0)) for item in rust_directory_files
            ),
        }
    rust_issue_build_time_ms = rust_report.get(
        "issue_build_time_ms",
        sum(item.get("issue_build_time_ms", 0) for item in rust_directory_files),
    )
    python_template_recommendation_summary = None
    if python_template_recommendation is not None:
        python_template_recommendation_summary = {
            "matched": python_template_recommendation["matched"],
            "confidence": python_template_recommendation["confidence"],
            "selected_template": (
                "redacted" if python_template_recommendation["template"] else None
            ),
            "alternative_count": len(python_template_recommendation["alternatives"]),
        }
    rust_template_recommendation_summary = None
    if rust_template_recommendation is not None:
        rust_template_recommendation_summary = {
            "engine": rust_template_recommendation["engine"],
            "source_workbooks": rust_template_recommendation["source_workbooks"],
            "matched": rust_template_recommendation["matched"],
            "confidence": rust_template_recommendation["confidence"],
            "selected_template": (
                "redacted" if rust_template_recommendation["selected_template"] else None
            ),
            "alternative_count": len(rust_template_recommendation["alternatives"]),
            "elapsed_ms": rust_template_recommendation["elapsed_ms"],
            "matches_expected_template": rust_template_recommendation[
                "matches_expected_template"
            ],
        }
    summary = {
        "schema_version": 1,
        "suite": "audit.s1_conditional_issue_extraction",
        "baseline": "Open-Data-Audit ConditionalFormatService(PYTHON)",
        "status": (
            "passed"
            if mismatch_count == 0
            and template_selection_matches_expected is not False
            and template_recommendation_matches_python is not False
            else "failed"
        ),
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
        "area_source": "template_defined_names" if args.template_workbook else "conditional_format_applies_to",
        "python_scan_areas": len(ranges),
        "python_structure_areas": len(structure_ranges),
        "rust_area_mapping": rust_area_mapping,
        "rust_source_discovery": rust_source_discovery,
        "rust_template_recommendation": rust_template_recommendation_summary,
        "python_template_recommendation": python_template_recommendation_summary,
        "template_recommendation_matches_python": template_recommendation_matches_python,
        "template_selection_matches_expected": template_selection_matches_expected,
        "python_unsupported_cells": python_stats.get("unsupported_cells", 0),
        "rust_diagnostics": rust_evaluation.get(
            "_diagnostics_count", len(rust_evaluation.get("diagnostics", []))
        ),
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
            "rust_issue_build_time_ms": rust_issue_build_time_ms,
            "rust_total_time_ms": rust_report.get(
                "total_elapsed_ms",
                rust_evaluation.get("total_time_ms", 0) + rust_issue_build_time_ms,
            ),
            "scope": "single real-workbook run; not a performance benchmark",
        },
        "verification": {
            "source_workbook": "read-only",
            "template_workbook": "read-only" if args.template_workbook else "not used",
            "mapping_config": "read-only" if args.mapping_config else "built-in defaults",
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
    return (
        0
        if mismatch_count == 0 and template_selection_matches_expected is not False
        and template_recommendation_matches_python is not False
        else 1
    )


if __name__ == "__main__":
    raise SystemExit(main())
