#!/usr/bin/env python3
"""Generate the read-only OpenPyXL baseline for S2 external workbook parsing."""

from __future__ import annotations

import hashlib
import json
import sys
from dataclasses import asdict
from decimal import Decimal
from pathlib import Path

sys.dont_write_bytecode = True

WORKSPACE = Path(__file__).resolve().parents[2]
DEFAULT_SOURCE = Path(r"C:\Users\AZI\Desktop\VibeCode\3.open-data-audit")
FIXTURES = WORKSPACE / "tests" / "fixtures" / "s2_external_import"
DEFAULT_OUTPUT = WORKSPACE / "probe" / "reports" / "s2_external_import_python_baseline.json"


def canonical(value: object) -> object:
    if isinstance(value, Decimal):
        parts = value.as_tuple()
        return {
            "type": "decimal",
            "value": {"sign": parts.sign, "digits": list(parts.digits), "exponent": parts.exponent},
        }
    if isinstance(value, tuple):
        return [canonical(item) for item in value]
    if isinstance(value, list):
        return [canonical(item) for item in value]
    if isinstance(value, dict):
        return {key: canonical(item) for key, item in value.items()}
    return value


def external_item_payload(item, source_file: Path) -> dict[str, object]:
    source = asdict(item.source)
    source["source_file"] = source_file.name
    return {
        "data_date": item.data_date,
        "row_label": item.row_label,
        "indicator_name": item.indicator_name,
        "value": canonical(item.value),
        "source": source,
        "fallback_reason": item.fallback_reason,
    }


def parsed_payload(values, org_map: dict[str, str], source_file: Path, fallback_key) -> dict[str, object]:
    ordinary = []
    missing_label_fallback = None
    for key, item in values.items():
        if key == fallback_key:
            missing_label_fallback = external_item_payload(item, source_file)
        else:
            ordinary.append(external_item_payload(item, source_file))
    ordinary.sort(key=lambda item: (item["data_date"], item["row_label"], item["indicator_name"]))
    return {
        "time_series": values.time_series,
        "values": ordinary,
        "available_dates": sorted(values.available_dates),
        "date_labels": sorted([list(item) for item in values.date_labels]),
        "row_labels": sorted(values.row_labels),
        "indicator_names": sorted(values.indicator_names),
        "blank_values": sorted([list(item) for item in values.blank_values]),
        "missing_label_fallback": missing_label_fallback,
        "warnings": list(values.warnings),
        "external_org_map": dict(sorted(org_map.items())),
    }


def main() -> int:
    import argparse

    parser = argparse.ArgumentParser()
    parser.add_argument("--source-root", type=Path, default=DEFAULT_SOURCE)
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    args = parser.parse_args()

    source_root = args.source_root.resolve()
    sys.path.insert(0, str(source_root / "core" / "src"))
    import openpyxl

    from base_audit.systems.s2_report_collection import importer

    cases = json.loads((FIXTURES / "cases.json").read_text(encoding="utf-8"))["cases"]
    results = {}
    for case in cases:
        path = FIXTURES / case["file"]
        try:
            values, org_map = importer.import_external_workbook(path, reader_engine="OPENPYXL")
            results[case["id"]] = {
                "result": parsed_payload(
                    values,
                    org_map,
                    path,
                    importer.MISSING_INSTITUTION_LABEL_FALLBACK_KEY,
                )
            }
        except Exception as error:
            results[case["id"]] = {"error": str(error)}

    source_dir = source_root / "core" / "src" / "base_audit" / "systems" / "s2_report_collection"
    source_file = source_dir / "importer.py"
    report = {
        "status": "python_baseline_generated",
        "source": {
            "project": str(source_root),
            "module": "base_audit.systems.s2_report_collection.importer.import_external_workbook",
            "reader_engine": "OPENPYXL",
            "sha256": {"importer.py": hashlib.sha256(source_file.read_bytes()).hexdigest().upper()},
            "baseline_python": sys.version.split()[0],
            "openpyxl": openpyxl.__version__,
        },
        "fixture": {"path": "tests/fixtures/s2_external_import/cases.json", "case_count": len(cases)},
        "result": results,
        "scope": "S2 external .xlsx workbook parsing: ordinary/time-series layouts, date and value normalization, blank-value tracking, warnings, organization reference sheet selection, and expected input errors. Does not cover service orchestration or report export.",
    }
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({"output": str(args.output), "case_count": len(cases), "source_sha256": report["source"]["sha256"]}, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
