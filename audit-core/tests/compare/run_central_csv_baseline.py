#!/usr/bin/env python3
"""Run the read-only Open-Data-Audit CSV importer over the Rust parity cases."""

from __future__ import annotations

import argparse
import dataclasses
import json
from pathlib import Path
import sys
from tempfile import TemporaryDirectory


def jsonable(value):
    if dataclasses.is_dataclass(value):
        return {field.name: jsonable(getattr(value, field.name)) for field in dataclasses.fields(value)}
    if isinstance(value, dict):
        return {key: jsonable(item) for key, item in value.items()}
    if isinstance(value, (list, tuple)):
        return [jsonable(item) for item in value]
    return value


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--source-root", required=True)
    parser.add_argument(
        "--cases",
        default=str(Path(__file__).resolve().parents[1] / "fixtures" / "central_csv_cases.json"),
    )
    parser.add_argument(
        "--output",
        default=str(Path(__file__).resolve().parents[2] / "probe" / "reports" / "central_csv_python_baseline.json"),
    )
    args = parser.parse_args()

    source_root = Path(args.source_root).resolve()
    python_source = source_root / "core" / "src"
    if not (python_source / "base_audit" / "systems" / "s3_central_statistics" / "csv_importer.py").is_file():
        python_source = source_root
    sys.dont_write_bytecode = True
    sys.path.insert(0, str(python_source))
    from base_audit.systems.s3_central_statistics.csv_importer import read_central_csv

    cases = json.loads(Path(args.cases).read_text(encoding="utf-8"))["cases"]
    results = []
    with TemporaryDirectory(prefix="audit-central-csv-") as temporary:
        for case in cases:
            path = Path(temporary) / case["file_name"]
            path.write_bytes(case["csv_text"].encode(case["encoding"]))
            dataset = read_central_csv(
                path,
                unit_factor=case["unit_factor"],
                exempt_indicators=set(case["exempt_indicators"]),
            )
            results.append({"id": case["id"], "dataset": jsonable(dataset)})

    report = {
        "suite": "audit.s3.central_csv_importer",
        "baseline": "Open-Data-Audit Python read_central_csv",
        "cases": results,
    }
    output = Path(args.output)
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(json.dumps(report, ensure_ascii=False, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    print(f"wrote {len(results)} Python baseline cases to {output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
