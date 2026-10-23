#!/usr/bin/env python3
"""Compare Rust S1 template-rule discovery with the read-only Python baseline."""

from __future__ import annotations

import argparse
import dataclasses
import hashlib
import json
import os
import subprocess
import sys
import time
from pathlib import Path
from typing import Any


sys.dont_write_bytecode = True


RULE_FIELDS = (
    "rule_id",
    "enabled",
    "report_code",
    "sheet_name",
    "formula_cell",
    "target_cell",
    "severity",
    "message",
    "copy_range",
    "result_mode",
    "value_from_result",
)
RANGE_FIELDS = ("sheet_name", "address")


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def wsl_path(path: Path) -> str:
    drive = path.drive.rstrip(":").lower()
    if len(drive) != 1:
        raise ValueError(f"无法转换为 WSL 挂载路径：{path.drive}")
    relative = path.as_posix().split("/", 1)[1]
    return f"/mnt/{drive}/{relative}"


def invoke_rust(args: argparse.Namespace, template: Path, config: Path | None) -> dict[str, Any]:
    executable = args.rust_executable.resolve(strict=True)
    if os.name == "nt" and executable.suffix.lower() != ".exe":
        command = [
            "wsl.exe",
            "--distribution",
            args.wsl_distribution,
            "--exec",
            wsl_path(executable),
        ]
        template_arg = wsl_path(template)
        config_arg = wsl_path(config) if config else None
    else:
        command = [str(executable)]
        template_arg = str(template)
        config_arg = str(config) if config else None

    command.extend(
        [
            "s1-template-rules",
            "--template-workbook",
            template_arg,
            "--include-rule-definitions",
        ]
    )
    if config_arg:
        command.extend(["--mapping-config", config_arg])
    started = time.perf_counter()
    result = subprocess.run(
        command,
        check=False,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )
    elapsed_ms = round((time.perf_counter() - started) * 1000)
    if result.returncode:
        raise RuntimeError(
            f"Rust S1 模板规则 CLI 返回 {result.returncode}：{result.stderr.strip()}"
        )
    payload = json.loads(result.stdout)
    if payload.get("status") != "completed" or not isinstance(payload.get("definition"), dict):
        raise ValueError("Rust CLI 没有返回完整模板规则定义")
    payload["_comparison_elapsed_ms"] = elapsed_ms
    return payload


def normalized_definition(definition: Any) -> dict[str, Any]:
    return json.loads(json.dumps(dataclasses.asdict(definition), ensure_ascii=False))


def compare_definitions(
    python_definition: dict[str, Any], rust_definition: dict[str, Any]
) -> dict[str, int]:
    differences: dict[str, int] = {}
    for field in ("structured",):
        differences[field] = int(python_definition.get(field) != rust_definition.get(field))

    python_rules = python_definition.get("rules", [])
    rust_rules = rust_definition.get("rules", [])
    differences["rule_count"] = abs(len(python_rules) - len(rust_rules))
    for field in RULE_FIELDS:
        count = sum(
            left.get(field) != right.get(field)
            for left, right in zip(python_rules, rust_rules)
        )
        differences[f"rules.{field}"] = count

    for field in ("copy_ranges", "structure_ranges", "extraction_ranges"):
        python_ranges = python_definition.get(field, [])
        rust_ranges = rust_definition.get(field, [])
        count = abs(len(python_ranges) - len(rust_ranges))
        count += sum(
            any(left.get(item) != right.get(item) for item in RANGE_FIELDS)
            for left, right in zip(python_ranges, rust_ranges)
        )
        differences[field] = count
    return differences


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--source-root", type=Path, required=True)
    parser.add_argument("--template-workbook", type=Path, required=True)
    parser.add_argument("--rust-executable", type=Path, required=True)
    parser.add_argument("--mapping-config", type=Path)
    parser.add_argument("--wsl-distribution", default="Ubuntu")
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()

    source_root = args.source_root.resolve(strict=True)
    template_path = args.template_workbook.resolve(strict=True)
    mapping_config = args.mapping_config.resolve(strict=True) if args.mapping_config else None
    before = sha256(template_path)
    sys.path.insert(0, str(source_root / "core" / "src"))
    from base_audit.name_config import (
        FORMULA_COPY_FUNCTION,
        ISSUE_EXTRACT_FUNCTION,
        STRUCTURE_COMPARE_FUNCTION,
        features_of_type,
        load_feature_mappings,
    )
    from base_audit.native.openpyxl_workbook import read_template

    python_started = time.perf_counter()
    mappings = load_feature_mappings(mapping_config, template_path)
    definition = read_template(
        template_path,
        formula_mappings=features_of_type(mappings, FORMULA_COPY_FUNCTION),
        structure_mappings=features_of_type(mappings, STRUCTURE_COMPARE_FUNCTION),
        extraction_mappings=features_of_type(mappings, ISSUE_EXTRACT_FUNCTION),
    )
    python_elapsed_ms = round((time.perf_counter() - python_started) * 1000)
    python_payload = normalized_definition(definition)

    rust_payload = invoke_rust(args, template_path, mapping_config)
    rust_definition = rust_payload["definition"]
    field_differences = compare_definitions(python_payload, rust_definition)
    mismatch_count = sum(field_differences.values())
    python_rules = python_payload["rules"]
    rust_summary = rust_payload["summary"]
    python_summary = {
        "template_mode": "structured" if python_payload["structured"] else "named_formula_ranges",
        "rules_total": len(python_rules),
        "rules_enabled": sum(bool(rule["enabled"]) for rule in python_rules),
        "formula_named_ranges": (
            0 if python_payload["structured"] else len(python_payload["copy_ranges"])
        ),
        "copy_ranges": len(python_payload["copy_ranges"]),
        "structure_ranges": len(python_payload["structure_ranges"]),
        "extraction_ranges": len(python_payload["extraction_ranges"]),
        "keyword_rules": sum(rule["result_mode"] == "keyword" for rule in python_rules),
        "error_only_rules": sum(rule["result_mode"] == "error_only" for rule in python_rules),
    }
    parity_fields = (
        "template_mode",
        "rules_total",
        "rules_enabled",
        "formula_named_ranges",
        "copy_ranges",
        "structure_ranges",
        "extraction_ranges",
        "keyword_rules",
        "error_only_rules",
    )
    summary_differences = {
        field: int(python_summary[field] != rust_summary.get(field))
        for field in parity_fields
    }
    mismatch_count += sum(summary_differences.values())
    after = sha256(template_path)
    report = {
        "schema_version": 1,
        "suite": "audit.s1_template_rule_discovery",
        "baseline": "Open-Data-Audit base_audit.native.openpyxl_workbook.read_template",
        "status": "passed" if mismatch_count == 0 else "mismatch",
        "sample": {
            "template_workbook": "redacted real S1 template",
            "mapping_config": "provided" if mapping_config else "source defaults",
        },
        "python": {
            **python_summary,
            "elapsed_ms": python_elapsed_ms,
        },
        "rust": {
            **{
                field: rust_summary.get(field)
                for field in (
                    *parity_fields,
                    "engine",
                    "formula_cells",
                    "rule_comments",
                    "explicit_comment_ids",
                    "read_time_ms",
                    "elapsed_ms",
                )
            },
            "cli_elapsed_ms": rust_payload["_comparison_elapsed_ms"],
        },
        "comparison": {
            "template_definition_field_difference_counts": field_differences,
            "summary_difference_counts": summary_differences,
            "mismatch_count": mismatch_count,
        },
        "verification": {
            "template_unchanged": before == after,
            "raw_rule_definitions_or_formula_text_persisted": False,
            "formula_recalculation": "not performed",
            "scope": "one real template; all discovered rules and mapped ranges compared in memory",
        },
    }
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(
        json.dumps(report, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0 if report["status"] == "passed" else 1


if __name__ == "__main__":
    raise SystemExit(main())
