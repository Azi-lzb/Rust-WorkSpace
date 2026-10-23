#!/usr/bin/env python3
"""Generate a read-only Python baseline for S2 report summary aggregation."""

from __future__ import annotations

import argparse
import hashlib
import json
import sys
from dataclasses import asdict
from pathlib import Path

sys.dont_write_bytecode = True

WORKSPACE = Path(__file__).resolve().parents[2]
DEFAULT_SOURCE = Path(r"C:\Users\AZI\Desktop\VibeCode\3.open-data-audit")
FIXTURE = WORKSPACE / "tests" / "fixtures" / "s2_summary" / "cases.json"
DEFAULT_OUTPUT = WORKSPACE / "probe" / "reports" / "s2_summary_python_baseline.json"


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--source-root", type=Path, default=DEFAULT_SOURCE)
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    args = parser.parse_args()

    source_root = args.source_root.resolve()
    sys.path.insert(0, str(source_root / "core" / "src"))
    from base_audit.systems.s2_report_collection import models
    from base_audit.systems.s2_report_collection.exporter import build_summary

    fixtures = json.loads(FIXTURE.read_text(encoding="utf-8"))["cases"]
    results = {}
    for case in fixtures:
        records = []
        for index, item in enumerate(case["records"], start=1):
            form_code = str(item["form_code"])
            records.append(
                models.IndicatorRecord(
                    institution_id=str(item["institution_id"]),
                    institution_name="",
                    social_credit_code="",
                    institution_type="",
                    handling_branch="",
                    region="",
                    form_code=form_code,
                    indicator_code="",
                    indicator_name="",
                    period="",
                    source_date="",
                    value=None,
                    value_type="",
                    source_unit="",
                    normalized_value=None,
                    normalized_unit="",
                    source=models.SourceRef("s2_summary_fixture.xlsx", form_code, index),
                )
            )
        findings = [
            models.AuditFinding(
                finding_id=f"summary-fixture-{index}",
                audit_type=str(item["audit_type"]),
                status=str(item["status"]),
                severity=str(item["severity"]),
            )
            for index, item in enumerate(case["findings"], start=1)
        ]
        results[case["id"]] = asdict(build_summary(records, findings))

    module_dir = source_root / "core" / "src" / "base_audit" / "systems" / "s2_report_collection"
    report = {
        "status": "python_baseline_generated",
        "source": {
            "module": "base_audit.systems.s2_report_collection.exporter.build_summary",
            "exporter_sha256": hashlib.sha256((module_dir / "exporter.py").read_bytes()).hexdigest().upper(),
            "models_sha256": hashlib.sha256((module_dir / "models.py").read_bytes()).hexdigest().upper(),
            "baseline_python": sys.version.split()[0],
        },
        "fixture": {"path": "tests/fixtures/s2_summary/cases.json", "case_count": len(fixtures)},
        "result": results,
        "scope": "S2 summary aggregation only: unique institution/form counts, record/finding totals, and abnormal finding type/severity counts.",
    }
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(report, ensure_ascii=False, indent=2, allow_nan=False) + "\n", encoding="utf-8")
    print(json.dumps({"output": str(args.output), "case_count": len(fixtures), **report["source"]}, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
