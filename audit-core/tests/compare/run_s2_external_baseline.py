#!/usr/bin/env python3
"""Generate a read-only Python baseline for the S2 external comparison slice."""

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
FIXTURE = WORKSPACE / "tests" / "fixtures" / "s2_external" / "cases.json"
DEFAULT_OUTPUT = WORKSPACE / "probe" / "reports" / "s2_external_python_baseline.json"


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
    value = decode_value(payload["value"])
    normalized = decode_value(payload["normalized_value"])
    source_payload = payload["source"]
    source = models.SourceRef(
        str(source_payload["source_file"]),
        str(source_payload["source_sheet"]),
        int(source_payload["source_row"]),
    )
    return models.IndicatorRecord(
        str(payload["institution_id"]),
        str(payload["institution_name"]),
        str(payload["social_credit_code"]),
        str(payload["institution_type"]),
        str(payload["handling_branch"]),
        str(payload["region"]),
        str(payload["form_code"]),
        str(payload["indicator_code"]),
        str(payload["indicator_name"]),
        str(payload["period"]),
        str(payload["source_date"]),
        value,
        str(payload["value_type"]),
        str(payload["source_unit"]),
        normalized,
        str(payload["normalized_unit"]),
        source,
    )


def external_from_payload(payload: dict[str, object], importer, models):
    external = importer.ExternalValues(time_series=bool(payload.get("time_series", False)))
    external.available_dates = set(payload.get("available_dates", []))
    external.date_labels = {tuple(item) for item in payload.get("date_labels", [])}
    external.row_labels = set(payload.get("row_labels", []))
    external.indicator_names = set(payload.get("indicator_names", []))
    for item in payload.get("values", []):
        source_payload = item["source"]
        source = models.SourceRef(
            str(source_payload["source_file"]),
            str(source_payload["source_sheet"]),
            int(source_payload["source_row"]),
        )
        value = importer.ExternalValue(
            str(item.get("row_label", "")),
            str(item.get("indicator_name", "")),
            decode_value(item["value"]),
            source,
            str(item.get("fallback_reason", "")),
            str(item.get("data_date", "")),
        )
        if item.get("data_date"):
            key = (str(item["data_date"]), value.row_label, value.indicator_name)
        else:
            key = (value.row_label, value.indicator_name)
        external[key] = value
    fallback = payload.get("missing_label_fallback")
    if fallback is not None:
        source_payload = fallback["source"]
        source = models.SourceRef(
            str(source_payload["source_file"]),
            str(source_payload["source_sheet"]),
            int(source_payload["source_row"]),
        )
        external[importer.MISSING_INSTITUTION_LABEL_FALLBACK_KEY] = importer.ExternalValue(
            str(fallback.get("row_label", "")),
            str(fallback.get("indicator_name", "")),
            decode_value(fallback["value"]),
            source,
            str(fallback.get("fallback_reason", "")),
            str(fallback.get("data_date", "")),
        )
    return external


def config_from_payload(payload: dict[str, object], models):
    indicators = {
        str(code): models.IndicatorConfig(
            str(code), data_type=str(item.get("data_type", ""))
        )
        for code, item in payload.get("indicators", {}).items()
    }
    institutions_by_id = {
        str(code): models.InstitutionProfile(
            str(code), str(item.get("institution_name", code)), report_item=str(item.get("report_item", ""))
        )
        for code, item in payload.get("institutions_by_id", {}).items()
    }
    institutions_by_name = {
        str(name): models.InstitutionProfile(
            str(item.get("institution_id", name)), str(name), report_item=str(item.get("report_item", ""))
        )
        for name, item in payload.get("institutions_by_name", {}).items()
    }
    rules = [
        models.ExternalCheckRule(
            str(item["rule_id"]),
            str(item["indicator_code"]),
            "大集中",
            str(item["external_indicator_name"]),
            str(item.get("comparison_method", "等于")),
            float(item.get("tolerance", 0.0)),
            str(item.get("severity", "硬性")),
            bool(item.get("enabled", True)),
        )
        for item in payload.get("external_rules", [])
    ]
    units = payload.get("unit_settings", {})
    return models.ReportAuditConfig(
        indicators=indicators,
        institutions_by_id=institutions_by_id,
        institutions_by_name=institutions_by_name,
        external_rules=rules,
        unit_settings=models.UnitSettings(
            "元", str(units.get("external_file_unit", "元")), str(units.get("output_file_unit", "元"))
        ),
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
    import argparse

    parser = argparse.ArgumentParser()
    parser.add_argument("--source-root", type=Path, default=DEFAULT_SOURCE)
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    args = parser.parse_args()

    source_root = args.source_root.resolve()
    sys.path.insert(0, str(source_root / "core" / "src"))
    from base_audit.systems.s2_report_collection import importer, models
    from base_audit.systems.s2_report_collection.external_engine import ExternalComparisonEngine

    fixtures = json.loads(FIXTURE.read_text(encoding="utf-8"))["cases"]
    results = {}
    for case in fixtures:
        current = [record_from_payload(item, models) for item in case["current"]]
        current_map = {record.logical_key: record for record in current}
        external = external_from_payload(case["external"], importer, models)
        config = config_from_payload(case["config"], models)
        try:
            findings = ExternalComparisonEngine().run(
                current_map,
                external,
                {str(key): str(value) for key, value in case.get("external_org_map", {}).items()},
                config,
            )
            results[case["id"]] = {"findings": [finding_payload(item) for item in findings]}
        except Exception as error:
            results[case["id"]] = {"error": str(error)}

    source_dir = source_root / "core" / "src" / "base_audit" / "systems" / "s2_report_collection"
    source_files = {
        name: source_dir / name
        for name in ("external_engine.py", "importer.py", "models.py", "unit_conversion.py")
    }
    report = {
        "status": "python_baseline_generated",
        "source": {
            "project": str(source_root),
            "module": "base_audit.systems.s2_report_collection.external_engine.ExternalComparisonEngine",
            "sha256": {name: hashlib.sha256(path.read_bytes()).hexdigest().upper() for name, path in source_files.items()},
            "baseline_python": sys.version.split()[0],
        },
        "fixture": {"path": "tests/fixtures/s2_external/cases.json", "case_count": len(fixtures)},
        "result": results,
        "scope": "S2 external comparison rules over imported records and external-data metadata. Does not cover external workbook parsing, configuration workbook loading, period comparison, expression rules, report export or service orchestration.",
    }
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(report, ensure_ascii=False, indent=2, allow_nan=False) + "\n", encoding="utf-8")
    print(json.dumps({
        "output": str(args.output),
        "case_count": len(fixtures),
        "finding_count": sum(len(item.get("findings", [])) for item in results.values()),
        "error_count": sum("error" in item for item in results.values()),
        "source_sha256": report["source"]["sha256"],
    }, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
