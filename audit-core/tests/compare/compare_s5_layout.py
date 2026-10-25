#!/usr/bin/env python3
"""Compare the Rust S5 layout-region port with the Python baseline.

For each real template workbook, both sides produce named regions, numeric
regions (with origins), inferred numeric regions (results + missing), and the
structure-region grouping; the comparator diffs every region field verbatim
and writes aggregate counts only.
"""

from __future__ import annotations

import argparse
import json
import os
import subprocess
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[3]


def python_side(workbook_path: Path):
    from openpyxl import load_workbook

    from s1audit_configured.layout import (
        inferred_numeric_regions,
        named_regions,
        numeric_regions,
    )

    workbook_path = load_workbook(workbook_path, data_only=False)

    def region_dict(region):
        return {
            "name": region.name,
            "sheet": region.sheet,
            "reference": region.reference,
            "min_row": region.min_row,
            "min_col": region.min_col,
            "max_row": region.max_row,
            "max_col": region.max_col,
        }

    named = [region_dict(region) for region in named_regions(workbook_path, ("表结构区域", "行索引区域", "列索引区域"))]
    numeric, origins = numeric_regions(workbook_path)
    inferred, missing = inferred_numeric_regions(workbook_path)
    structure: dict[str, list] = {}
    for region in named:
        structure.setdefault(region["sheet"], []).append(region)
    for regions in structure.values():
        from s1audit_configured.layout import region_number as _rn

        regions.sort(key=lambda item: (_rn(item["name"]), item["min_col"], item["min_row"], item["name"]))
    return {
        "named_structure_axes": named,
        "numeric_regions": [region_dict(region) for region in numeric],
        "numeric_origins": origins,
        "inferred_results": [region_dict(region) for region in inferred],
        "inferred_missing": [region_dict(region) for region in missing],
        "structure_by_sheet": structure,
    }


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
    parser.add_argument("--workbooks", type=Path, nargs="+", required=True,
                        help="real template/config workbooks (read-only)")
    parser.add_argument("--workspace", type=Path, default=ROOT)
    parser.add_argument("--report", type=Path, required=True)
    args = parser.parse_args()

    source_root = args.source_root.resolve(strict=True)
    workspace = args.workspace.resolve(strict=True)
    report_path = args.report.resolve()
    sys.path.insert(0, str(source_root / "core" / "src"))

    comparisons = []
    total_diffs = 0
    with tempfile.TemporaryDirectory(prefix="audit-s5-layout-") as temporary:
        directory = Path(temporary)
        for source_path in args.workbooks:
            source_path = source_path.resolve(strict=True)
            work_copy = directory / source_path.name
            work_copy.write_bytes(source_path.read_bytes())

            expected = python_side(work_copy)

            rust_request = json.dumps({"workbook": str(work_copy)}, ensure_ascii=False)
            request_path = directory / f"request_{len(comparisons)}.json"
            request_path.write_text(rust_request, encoding="utf-8")
            rust_report = directory / f"rust_{len(comparisons)}.json"
            environment = os.environ.copy()
            environment["AUDIT_S5_LAYOUT_INPUT"] = str(request_path)
            environment["AUDIT_S5_LAYOUT_REPORT"] = str(rust_report)
            completed = subprocess.run(
                [
                    "cargo", "test", "-p", "audit-core", "--test", "s5_layout_parity",
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
                    "workbook_index": len(comparisons),
                    "exit_code": completed.returncode,
                    "stderr_tail": completed.stderr[-1500:],
                }))
                return 1
            actual = json.loads(rust_report.read_text(encoding="utf-8"))

            diffs: list[str] = []
            diff(expected, actual, "layout", diffs)
            total_diffs += len(diffs)
            comparisons.append({
                "workbook_index": len(comparisons),
                "python_named_structure_axes": len(expected["named_structure_axes"]),
                "python_numeric_regions": len(expected["numeric_regions"]),
                "python_numeric_origins": expected["numeric_origins"],
                "python_inferred_results": len(expected["inferred_results"]),
                "python_inferred_missing": len(expected["inferred_missing"]),
                "rust_status": actual.get("status", "unknown"),
                "diff_count": len(diffs),
                "diff_paths_sample": sorted(set(diffs))[:15],
            })

    report = {
        "schema_version": 1,
        "suite": "audit.s5_layout_regions_parity",
        "status": "completed" if total_diffs == 0 else "mismatch",
        "sample_scope": {
            "workbooks": len(comparisons),
            "workbooks_copied_to_temporary": True,
        },
        "comparison": {
            "total_diff_count": total_diffs,
            "cases": comparisons,
        },
        "verification": {
            "rust_runner": "audit-core/tests/s5_layout_parity.rs",
            "python_baseline": "s1audit_configured.layout region discovery",
            "inputs_unchanged": True,
            "raw_sheet_names_persisted": False,
            "capability_status": "unverified",
            "production_route_eligible": False,
        },
        "limitations": [
            "Covers the defined-name region discovery half of layout.py; the "
            "sequence matching inside scan_rule_paths belongs to the "
            "address-scan slice.",
            "Sheet-scoped defined names are read through the shared "
            "read_defined_names records; multi-area names use the Python "
            "#index suffix convention.",
        ],
    }
    report_path.parent.mkdir(parents=True, exist_ok=True)
    report_path.write_text(
        json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )
    print(json.dumps({
        "status": report["status"],
        "workbooks": len(comparisons),
        "total_diffs": total_diffs,
    }))
    return 0 if total_diffs == 0 else 1


if __name__ == "__main__":
    import os

    raise SystemExit(main())
