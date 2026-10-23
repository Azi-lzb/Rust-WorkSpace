#!/usr/bin/env python3
"""Generate a read-only Python baseline for the S2 rule-engine slice."""

from __future__ import annotations

import argparse
import hashlib
import json
import sys
from dataclasses import asdict
from decimal import Decimal
from pathlib import Path

sys.dont_write_bytecode = True

WORKSPACE = Path(__file__).resolve().parents[2]
DEFAULT_SOURCE = Path(r"C:\Users\AZI\Desktop\VibeCode\3.open-data-audit")
FIXTURE = WORKSPACE / "tests" / "fixtures" / "s2_rule_engine" / "cases.json"
DEFAULT_OUTPUT = WORKSPACE / "probe" / "reports" / "s2_rule_engine_python_baseline.json"


def decode_value(payload: dict[str, object]) -> object:
    kind = payload["type"]
    value = payload.get("value")
    if kind == "empty":
        return None
    if kind == "decimal":
        assert isinstance(value, dict)
        return Decimal((int(value["sign"]), tuple(value["digits"]), int(value["exponent"])))
    if kind == "number":
        return float(value)
    if kind == "integer":
        return int(value)
    if kind == "boolean":
        return bool(value)
    if kind == "string":
        return str(value)
    raise ValueError(f"unsupported fixture value kind: {kind}")


def record_from_payload(payload: dict[str, object], models):
    code = str(payload["indicator_code"])
    period = str(payload["period"])
    form_code = str(payload.get("form_code", "20202"))
    normalized = decode_value(payload["normalized_value"])
    source = models.SourceRef(
        str(payload.get("source_file", "synthetic.xlsx")),
        str(payload.get("source_sheet", form_code)),
        int(payload.get("source_row", 1)),
    )
    return models.IndicatorRecord(
        str(payload.get("institution_id", "9144")),
        str(payload.get("institution_name", "甲银行")),
        str(payload.get("social_credit_code", "C9144")),
        str(payload.get("institution_type", "农商行")),
        str(payload.get("handling_branch", "一部")),
        str(payload.get("region", "示例地区")),
        form_code,
        code,
        str(payload.get("indicator_name", code)),
        period,
        str(payload.get("source_date", f"{period}-30")),
        normalized,
        str(payload.get("value_type", "数值")),
        str(payload.get("source_unit", "元")),
        normalized,
        str(payload.get("normalized_unit", "万元")),
        source,
    )


def config_from_payload(payload: dict[str, object], models):
    indicators = {
        str(code): models.IndicatorConfig(
            str(code),
            data_type=str(item.get("data_type", "")),
            disabled=bool(item.get("disabled", False)),
        )
        for code, item in payload.get("indicators", {}).items()
    }
    rules = [
        models.RuleConfig(
            rule_id=str(item["rule_id"]),
            rule_type=str(item["rule_type"]),
            description=str(item["description"]),
            expression=str(item["expression"]),
            severity=str(item.get("severity", "软性")),
            threshold=item.get("threshold"),
            invert=bool(item.get("invert", False)),
            disabled=bool(item.get("disabled", False)),
            note=str(item.get("note", "")),
            source=(
                models.SourceRef(
                    str(item["source"]["source_file"]),
                    str(item["source"]["source_sheet"]),
                    int(item["source"]["source_row"]),
                )
                if item.get("source")
                else None
            ),
        )
        for item in payload.get("rules", [])
    ]
    settings = payload.get("unit_settings", {})
    unit_settings = models.UnitSettings(
        source_unit=str(settings.get("source_unit", "元")),
        external_file_unit=str(settings.get("external_file_unit", "元")),
        output_file_unit=str(settings.get("output_file_unit", "元")),
    )
    return models.ReportAuditConfig(indicators=indicators, rules=rules, unit_settings=unit_settings)


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


def finding_payload(finding) -> dict[str, object]:
    payload = canonical(asdict(finding))
    for source in payload["source_refs"]:
        source["source_file"] = Path(source["source_file"]).name
    for related in payload["related_records"]:
        related["source"]["source_file"] = Path(related["source"]["source_file"]).name
    return payload


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--source-root", type=Path, default=DEFAULT_SOURCE)
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    args = parser.parse_args()

    source_root = args.source_root.resolve()
    sys.path.insert(0, str(source_root / "core" / "src"))
    from base_audit.systems.s2_report_collection import models
    from base_audit.systems.s2_report_collection.rule_engine import RuleEngine

    fixtures = json.loads(FIXTURE.read_text(encoding="utf-8"))["cases"]
    results = {}
    for case in fixtures:
        current_records = [record_from_payload(item, models) for item in case["current"]]
        previous_records = [record_from_payload(item, models) for item in case["previous"]]
        current = {record.logical_key: record for record in current_records}
        previous = {record.logical_key: record for record in previous_records}
        config = config_from_payload(case["config"], models)
        findings, quality = RuleEngine().run(current, previous, config)
        results[case["id"]] = {
            "findings": [finding_payload(item) for item in findings],
            "quality": [finding_payload(item) for item in quality],
        }

    base = source_root / "core" / "src" / "base_audit" / "systems" / "s2_report_collection"
    sources = {
        name: hashlib.sha256((base / name).read_bytes()).hexdigest().upper()
        for name in ("rule_engine.py", "models.py", "unit_conversion.py")
    }
    report = {
        "status": "python_baseline_generated",
        "source": {
            "project": str(source_root),
            "module": "base_audit.systems.s2_report_collection.rule_engine.RuleEngine",
            "source_sha256": sources,
            "baseline_python": sys.version.split()[0],
        },
        "fixture": {"path": "tests/fixtures/s2_rule_engine/cases.json", "case_count": len(fixtures)},
        "result": results,
        "scope": "S2 cumulative and expression rule evaluation over already imported records and in-memory rule configuration. Does not cover workbook loading, configuration parsing, import, external comparisons, report export or service orchestration.",
    }
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(report, ensure_ascii=False, indent=2, allow_nan=False) + "\n", encoding="utf-8")
    print(json.dumps({
        "output": str(args.output),
        "case_count": len(fixtures),
        "finding_count": sum(len(item["findings"]) for item in results.values()),
        "quality_count": sum(len(item["quality"]) for item in results.values()),
        "source_sha256": sources,
    }, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
