#!/usr/bin/env python3
"""Run the read-only Open-Data-Audit period-pair preview baseline."""
from __future__ import annotations

import argparse
import hashlib
import json
import shutil
import sys
from pathlib import Path
from tempfile import TemporaryDirectory


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--source-root", required=True)
    parser.add_argument(
        "--workspace-root",
        default=str(Path(__file__).resolve().parents[2]),
    )
    parser.add_argument(
        "--cases",
        default=str(
            Path(__file__).resolve().parents[1]
            / "fixtures"
            / "period_pairs_cases.json"
        ),
    )
    parser.add_argument(
        "--output",
        default=str(
            Path(__file__).resolve().parents[2]
            / "probe"
            / "reports"
            / "period_pairs_python_baseline.json"
        ),
    )
    args = parser.parse_args()
    source_root = Path(args.source_root).resolve()
    workspace_root = Path(args.workspace_root).resolve()
    source_python = source_root / "core" / "src"
    if not (source_python / "base_audit" / "report_source.py").is_file():
        source_python = source_root
    sys.dont_write_bytecode = True
    sys.path.insert(0, str(source_python))
    from base_audit.period_compare import list_period_pairs
    from base_audit.report_source import report_source_warning

    case_data = json.loads(Path(args.cases).read_text(encoding="utf-8"))
    fixture_root = workspace_root / "tests" / "fixtures"
    results = []
    for case in case_data["cases"]:
        with TemporaryDirectory(prefix="audit-period-pairs-") as temporary:
            root = Path(temporary)
            current = root / "current"
            previous = root / "previous"
            current.mkdir()
            previous.mkdir()
            for side, directory in (("current", current), ("previous", previous)):
                for item in case[side]:
                    shutil.copyfile(
                        fixture_root / item["fixture"],
                        directory / item["name"],
                    )
            pairs = list_period_pairs(current, previous)
            for pair in pairs:
                for key in ("curPath", "prePath"):
                    if pair[key]:
                        pair[key] = Path(pair[key]).name
            results.append(
                {
                    "id": case["id"],
                    "warning": report_source_warning(current, previous),
                    "pairs": pairs,
                }
            )

    report = {
        "schema_version": 1,
        "suite": case_data["suite"],
        "baseline": "base_audit.period_compare.list_period_pairs + report_source.report_source_warning",
        "source_python_version": sys.version.split()[0],
        "source_root": str(source_root),
        "source_sha256": {
            "period_compare.py": sha256(source_python / "base_audit" / "period_compare.py"),
            "report_source.py": sha256(source_python / "base_audit" / "report_source.py"),
        },
        "fixture_cases": len(case_data["cases"]),
        "results": results,
    }
    output = Path(args.output)
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(
        json.dumps(report, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    print(
        json.dumps(
            {"suite": report["suite"], "cases": len(results)},
            ensure_ascii=False,
        )
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
