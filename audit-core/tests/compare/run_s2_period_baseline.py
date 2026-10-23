#!/usr/bin/env python3
"""Generate a read-only Python baseline for the Rust S2 period engine slice."""

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
FIXTURE = WORKSPACE / "tests" / "fixtures" / "s2_period" / "cases.json"
DEFAULT_OUTPUT = WORKSPACE / "probe" / "reports" / "s2_period_python_baseline.json"


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
            period_enabled=bool(item.get("period_enabled", True)),
            period_strategy_group=str(item.get("period_strategy_group", "")),
            min_change_value=item.get("min_change_value"),
        )
        for code, item in payload.get("indicators", {}).items()
    }
    bands = [
        models.PeriodBandConfig(
            str(item.get("rule_id", "")),
            item.get("lower"),
            item.get("upper"),
            str(item.get("description", "")),
            str(item.get("severity", "")),
            bool(item.get("disabled", False)),
            str(item.get("strategy_group", "")),
            str(item.get("data_type", "")),
        )
        for item in payload.get("bands", [])
    ]
    return models.ReportAuditConfig(
        indicators=indicators,
        bands=bands,
        text_change_severity=str(payload.get("text_change_severity", "硬性")),
        current_missing_severity=str(payload.get("current_missing_severity", "硬性")),
        previous_missing_severity=str(payload.get("previous_missing_severity", "硬性")),
    )


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
    return payload


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--source-root", type=Path, default=DEFAULT_SOURCE)
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    args = parser.parse_args()

    source_root = args.source_root.resolve()
    sys.path.insert(0, str(source_root / "core" / "src"))
    from base_audit.systems.s2_report_collection import models
    from base_audit.systems.s2_report_collection.period_engine import (
        PeriodComparisonEngine,
        PeriodRuleExecutionError,
    )

    fixtures = json.loads(FIXTURE.read_text(encoding="utf-8"))["cases"]
    results = {}
    for case in fixtures:
        current_records = [record_from_payload(item, models) for item in case["current"]]
        previous_records = [record_from_payload(item, models) for item in case["previous"]]
        current = {record.logical_key: record for record in current_records}
        previous = {record.logical_key: record for record in previous_records}
        config = config_from_payload(case["config"], models)
        try:
            findings = PeriodComparisonEngine().run(current, previous, config)
            results[case["id"]] = {"findings": [finding_payload(item) for item in findings]}
        except PeriodRuleExecutionError as error:
            results[case["id"]] = {"error": str(error)}

    period_source = source_root / "core" / "src" / "base_audit" / "systems" / "s2_report_collection" / "period_engine.py"
    model_source = source_root / "core" / "src" / "base_audit" / "systems" / "s2_report_collection" / "models.py"
    report = {
        "status": "python_baseline_generated",
        "source": {
            "project": str(source_root),
            "module": "base_audit.systems.s2_report_collection.period_engine.PeriodComparisonEngine",
            "period_engine_sha256": hashlib.sha256(period_source.read_bytes()).hexdigest().upper(),
            "models_sha256": hashlib.sha256(model_source.read_bytes()).hexdigest().upper(),
            "baseline_python": sys.version.split()[0],
        },
        "fixture": {"path": "tests/fixtures/s2_period/cases.json", "case_count": len(fixtures)},
        "result": results,
        "scope": "S2 two-period comparison rules only: numeric/text/missing-side handling, band routing, decimal arithmetic, severity, findings and calculation traces. Does not cover configuration workbook loading, import, external comparison, rule expressions, report export or service orchestration.",
    }
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(report, ensure_ascii=False, indent=2, allow_nan=False) + "\n", encoding="utf-8")
    print(json.dumps({"output": str(args.output), "case_count": len(fixtures), "finding_count": sum(len(item.get("findings", [])) for item in results.values()), "error_count": sum("error" in item for item in results.values()), "period_engine_sha256": report["source"]["period_engine_sha256"], "models_sha256": report["source"]["models_sha256"]}, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
