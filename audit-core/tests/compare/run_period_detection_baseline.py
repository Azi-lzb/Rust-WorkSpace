#!/usr/bin/env python3
"""Run the original Open-Data-Audit period detection over synthetic trees."""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from tempfile import TemporaryDirectory


def scan_depth(value):
    if isinstance(value, bool):
        return value
    return int(value)


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--source-root", required=True)
    parser.add_argument(
        "--cases",
        default=str(
            Path(__file__).resolve().parents[1]
            / "fixtures"
            / "period_detection_cases.json"
        ),
    )
    parser.add_argument(
        "--output",
        default=str(
            Path(__file__).resolve().parents[2]
            / "probe"
            / "reports"
            / "period_detection_python_baseline.json"
        ),
    )
    args = parser.parse_args()
    source_root = Path(args.source_root).resolve()
    sys.dont_write_bytecode = True
    python_source = source_root / "core" / "src"
    if not (python_source / "base_audit" / "discovery.py").is_file():
        python_source = source_root
    sys.path.insert(0, str(python_source))
    from base_audit.discovery import detect_period, source_workbooks

    case_data = json.loads(Path(args.cases).read_text(encoding="utf-8"))
    results = []
    for case in case_data["cases"]:
        with TemporaryDirectory(prefix="audit-period-detection-") as temporary:
            cursor = Path(temporary)
            for name in case["parents"]:
                cursor = cursor / name
                cursor.mkdir()
            input_dir = cursor / case["root_name"]
            input_dir.mkdir()
            for directory in case["directories"]:
                (input_dir / directory).mkdir(parents=True, exist_ok=True)
            for name in case["files"]:
                path = input_dir / name
                path.parent.mkdir(parents=True, exist_ok=True)
                path.touch()

            depth = scan_depth(case["recursive"])
            files = source_workbooks(input_dir, recursive=depth)
            detected = detect_period(input_dir, recursive=depth)
            results.append(
                {
                    "id": case["id"],
                    "source_workbooks": sorted(
                        path.relative_to(input_dir).as_posix() for path in files
                    ),
                    "result": {
                        "period": detected.period,
                        "source": detected.source,
                        "conflict": detected.conflict,
                        "details": detected.details,
                    },
                }
            )

    report = {
        "schema_version": 1,
        "suite": "audit.period_detection",
        "baseline": "base_audit.discovery.source_workbooks + detect_period",
        "source_python_version": sys.version.split()[0],
        "source_root": str(source_root),
        "cases": results,
    }
    output = Path(args.output)
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({"suite": report["suite"], "cases": len(results)}, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
