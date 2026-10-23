#!/usr/bin/env python3
"""Generate the read-only Python parity baseline for S2 config checks."""

from __future__ import annotations

import argparse
import hashlib
import json
import sys
from pathlib import Path

import openpyxl


def normalize_text(value: str, path: Path) -> str:
    filename = path.name
    return value.replace(str(path), filename).replace(path.as_posix(), filename)


def normalize_report(report: dict, path: Path) -> dict:
    normalized = json.loads(json.dumps(report, ensure_ascii=False))
    normalized["path"] = path.name
    for issue in normalized.get("issues", []):
        issue["message"] = normalize_text(issue["message"], path)
    return normalized


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--source-root", required=True, type=Path)
    parser.add_argument(
        "--workspace-root", type=Path, default=Path(__file__).resolve().parents[3]
    )
    parser.add_argument(
        "--output", type=Path,
        default=Path(__file__).resolve().parents[2] / "probe" / "reports" / "s2_config_check_python_baseline.json",
    )
    args = parser.parse_args()

    source_root = args.source_root.resolve()
    sys.path.insert(0, str(source_root / "core" / "src"))
    from base_audit.systems.s2_report_collection.config_check import (  # noqa: PLC0415
        check_s2_config,
        format_config_report,
    )

    audit_root = args.workspace_root.resolve() / "audit-core"
    fixture_dir = audit_root / "tests" / "fixtures" / "s2_config"
    cases = json.loads((fixture_dir / "check_cases.json").read_text(encoding="utf-8"))["cases"]
    results: dict[str, dict] = {}
    for case in cases:
        path = fixture_dir / case["file"]
        report = check_s2_config(path)
        normalized = normalize_report(report, path)
        formatted = normalize_text(format_config_report(report), path)
        results[case["id"]] = {"report": normalized, "formatted": formatted}

    source_dir = source_root / "core" / "src" / "base_audit" / "systems" / "s2_report_collection"
    source_files = ("config_check.py", "config.py", "models.py", "unit_conversion.py")
    source_hashes = {
        name: hashlib.sha256((source_dir / name).read_bytes()).hexdigest().upper()
        for name in source_files
    }
    output = {
        "status": "python_baseline_generated",
        "source": {
            "project": "Open-Data-Audit",
            "module": "base_audit.systems.s2_report_collection.config_check.{check_s2_config,format_config_report}",
            "source_sha256": source_hashes,
            "python": sys.version.split()[0],
            "openpyxl": openpyxl.__version__,
        },
        "fixture": {"path": "tests/fixtures/s2_config/check_cases.json", "case_count": len(cases)},
        "result": results,
        "scope": "Read-only S2 configuration structure and runtime-parser checks over synthetic XLSX fixtures. Does not cover real business configurations, default workbook generation, service orchestration, report writing, or target Linux validation.",
    }
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(output, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({
        "output": str(args.output),
        "case_count": len(cases),
        "passed_cases": sum(item["report"]["passed"] for item in results.values()),
        "blocking_errors": sum(item["report"]["blocking_errors"] for item in results.values()),
        "warnings": sum(item["report"]["warnings"] for item in results.values()),
        "source_sha256": source_hashes,
    }, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
