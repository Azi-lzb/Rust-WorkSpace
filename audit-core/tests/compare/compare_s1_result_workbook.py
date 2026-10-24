#!/usr/bin/env python3
"""Compare the Rust S1-F01 result workbook with the read-only Python writer."""

from __future__ import annotations

import argparse
import json
import os
import sys
import tempfile
from pathlib import Path
from types import SimpleNamespace

sys.dont_write_bytecode = True
sys.path.insert(0, str(Path(__file__).resolve().parent))
from compare_s1_audit_directory import compare_result_workbooks


def source_path_for_python(value: str) -> str:
    if os.name == "nt" and value.lower().startswith("/mnt/") and len(value) > 6:
        drive = value[5].upper()
        return f"{drive}:\\{value[7:].replace('/', chr(92))}"
    return value


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--source-root", required=True, type=Path)
    parser.add_argument("--issues-json", required=True, type=Path)
    parser.add_argument("--rust-workbook", required=True, type=Path)
    args = parser.parse_args()

    if not args.rust_workbook.is_file():
        raise FileNotFoundError(f"Rust result workbook not found: {args.rust_workbook}")
    issues_data = json.loads(args.issues_json.read_text(encoding="utf-8"))
    issues = [SimpleNamespace(**item) for item in issues_data["issues"]]
    for issue in issues:
        issue.audit_file = source_path_for_python(issue.audit_file)

    sys.path.insert(0, str(args.source_root.resolve(strict=True) / "core" / "src"))
    from base_audit.native.result_writer import write_audit_summary

    with tempfile.TemporaryDirectory(prefix="audit-s1-result-parity-") as directory:
        python_workbook = Path(directory) / "python-result.xlsx"
        write_audit_summary(
            python_workbook,
            issues,
            sheet_name="本期审核结果",
            mode="OPENPYXL",
        )
        comparison = compare_result_workbooks(
            python_workbook,
            args.rust_workbook.resolve(strict=True),
        )

    result = {
        "status": comparison["status"],
        "issues": len(issues),
        "rows": comparison["rows"],
        "hyperlinks": comparison["hyperlinks"],
        "mismatch_counts": comparison["mismatch_counts"],
        "raw_values_or_paths_persisted": False,
    }
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0 if result["status"] == "passed" else 1


if __name__ == "__main__":
    raise SystemExit(main())
