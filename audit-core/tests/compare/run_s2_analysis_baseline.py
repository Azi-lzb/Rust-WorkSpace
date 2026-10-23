#!/usr/bin/env python3
"""Generate a read-only baseline for S2 in-memory service orchestration."""

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
FIXTURE = WORKSPACE / "tests" / "fixtures" / "s2_analysis" / "cases.json"
DEFAULT_OUTPUT = WORKSPACE / "probe" / "reports" / "s2_analysis_python_baseline.json"


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
            indicator_name=str(item.get("indicator_name", "")),
            form_code=str(item.get("form_code", "")),
            data_type=str(item.get("data_type", "")),
            disabled=bool(item.get("disabled", False)),
            note=str(item.get("note", "")),
            period_enabled=bool(item.get("period_enabled", True)),
            period_strategy_group=str(item.get("period_strategy_group", "")),
            min_change_value=item.get("min_change_value"),
        )
        for code, item in payload.get("indicators", {}).items()
    }
    institutions_by_id = {
        str(code): models.InstitutionProfile(
            institution_id=str(item.get("institution_id", code)),
            institution_name=str(item.get("institution_name", "")),
            social_credit_code=str(item.get("social_credit_code", "")),
            institution_type=str(item.get("institution_type", "")),
            handling_branch=str(item.get("handling_branch", "")),
            region=str(item.get("region", "")),
            report_item=str(item.get("report_item", "")),
        )
        for code, item in payload.get("institutions_by_id", {}).items()
    }
    institutions_by_name = {
        str(name): models.InstitutionProfile(
            institution_id=str(item.get("institution_id", "")),
            institution_name=str(item.get("institution_name", name)),
            social_credit_code=str(item.get("social_credit_code", "")),
            institution_type=str(item.get("institution_type", "")),
            handling_branch=str(item.get("handling_branch", "")),
            region=str(item.get("region", "")),
            report_item=str(item.get("report_item", "")),
        )
        for name, item in payload.get("institutions_by_name", {}).items()
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
    external_rules = [
        models.ExternalCheckRule(
            rule_id=str(item.get("rule_id", "")),
            indicator_code=str(item.get("indicator_code", "")),
            external_system=str(item.get("external_system", "大集中")),
            external_indicator_name=str(item.get("external_indicator_name", "")),
            comparison_method=str(item.get("comparison_method", "等于")),
            tolerance=float(item.get("tolerance", 0.0)),
            severity=str(item.get("severity", "硬性")),
            enabled=bool(item.get("enabled", True)),
        )
        for item in payload.get("external_rules", [])
    ]
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
        )
        for item in payload.get("rules", [])
    ]
    unit = payload.get("unit_settings", {})
    return models.ReportAuditConfig(
        indicators=indicators,
        institutions_by_id=institutions_by_id,
        institutions_by_name=institutions_by_name,
        bands=bands,
        external_rules=external_rules,
        rules=rules,
        warnings=[str(item) for item in payload.get("warnings", [])],
        unit_settings=models.UnitSettings(
            source_unit=str(unit.get("source_unit", "元")),
            external_file_unit=str(unit.get("external_file_unit", "元")),
            output_file_unit=str(unit.get("output_file_unit", "元")),
        ),
        hide_unchanged_period_rows=bool(payload.get("hide_unchanged_period_rows", False)),
        hide_matching_external_rows=bool(payload.get("hide_matching_external_rows", False)),
        text_change_severity=str(payload.get("text_change_severity", "硬性")),
        current_missing_severity=str(payload.get("current_missing_severity", "硬性")),
        previous_missing_severity=str(payload.get("previous_missing_severity", "硬性")),
    )


def import_result_from_payload(payload: dict[str, object], importer, models):
    records = [record_from_payload(item, models) for item in payload]
    return records


def external_from_payload(payload: dict[str, object], importer, models):
    data = payload["data"]
    external = importer.ExternalValues(time_series=bool(data.get("time_series", False)))
    external.available_dates = set(data.get("available_dates", []))
    external.date_labels = {tuple(item) for item in data.get("date_labels", [])}
    external.row_labels = set(data.get("row_labels", []))
    external.indicator_names = set(data.get("indicator_names", []))
    external.blank_values = {tuple(item) for item in data.get("blank_values", [])}
    external.warnings = [str(item) for item in payload.get("warnings", [])]
    for item in data.get("values", []):
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
    return external


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


def record_payload(record) -> dict[str, object]:
    payload = canonical(asdict(record))
    payload["value"] = imported_value_payload(record.value)
    payload["normalized_value"] = imported_value_payload(record.normalized_value)
    payload["source"]["source_file"] = Path(payload["source"]["source_file"]).name
    return payload


def imported_value_payload(value: object) -> dict[str, object]:
    if value is None:
        return {"type": "empty"}
    if isinstance(value, bool):
        return {"type": "boolean", "value": value}
    if isinstance(value, Decimal):
        return canonical(value)
    if isinstance(value, int):
        return {"type": "integer", "value": value}
    if isinstance(value, float):
        return {"type": "number", "value": value}
    if isinstance(value, str):
        return {"type": "string", "value": value}
    raise TypeError(f"unsupported imported record value: {type(value).__name__}")


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--source-root", type=Path, default=DEFAULT_SOURCE)
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    args = parser.parse_args()

    source_root = args.source_root.resolve()
    sys.path.insert(0, str(source_root / "core" / "src"))
    from base_audit.systems.s2_report_collection import importer, models
    from base_audit.systems.s2_report_collection.exporter import build_summary
    from base_audit.systems.s2_report_collection.external_engine import ExternalComparisonEngine
    from base_audit.systems.s2_report_collection.period_engine import PeriodComparisonEngine
    from base_audit.systems.s2_report_collection.rule_engine import RuleEngine
    from base_audit.systems.s2_report_collection.validation import (
        validate_config_warnings,
        validate_external_warnings,
        validate_import,
    )

    fixture = json.loads(FIXTURE.read_text(encoding="utf-8"))
    config = config_from_payload(fixture["config"], models)
    results: dict[str, object] = {}
    for case in fixture["cases"]:
        current_records = import_result_from_payload(case.get("current", []), importer, models)
        previous_records = import_result_from_payload(case.get("previous", []), importer, models)
        current_duplicates = [
            (
                record_from_payload(pair["first"], models),
                record_from_payload(pair["duplicate"], models),
            )
            for pair in case.get("current_duplicates", [])
        ]
        current_import = importer.ImportResult(
            records=current_records,
            duplicates=current_duplicates,
            warnings=[str(item) for item in case.get("current_warnings", [])],
        )
        previous_import = importer.ImportResult(
            records=previous_records,
            warnings=[str(item) for item in case.get("previous_warnings", [])],
        )
        current = current_import.by_logical_key()
        previous = previous_import.by_logical_key()

        quality = [
            *validate_config_warnings(config.warnings),
            *validate_import(current_import),
            *validate_import(previous_import),
        ]
        period_findings = PeriodComparisonEngine().run(current, previous, config)
        rule_findings, rule_quality = RuleEngine().run(current, previous, config)
        external_findings = []
        external = None
        external_org_map: dict[str, str] = {}
        if case.get("external") is not None:
            external = external_from_payload(case["external"], importer, models)
            external_org_map = {
                str(key): str(value)
                for key, value in case["external"].get("external_org_map", {}).items()
            }
            quality.extend(validate_external_warnings(external.warnings))
            external_findings = ExternalComparisonEngine().run(
                current, external, external_org_map, config
            )

        findings = [*period_findings, *external_findings, *rule_findings, *quality, *rule_quality]
        summary = build_summary(list(current.values()), findings)
        results[str(case["id"])] = {
            "current_records": [record_payload(item) for item in current.values()],
            "previous_records": [record_payload(item) for item in previous.values()],
            "period_findings": [finding_payload(item) for item in period_findings],
            "external_findings": [finding_payload(item) for item in external_findings],
            "rule_findings": [finding_payload(item) for item in rule_findings],
            "quality_findings": [finding_payload(item) for item in quality],
            "rule_quality_findings": [finding_payload(item) for item in rule_quality],
            "findings": [finding_payload(item) for item in findings],
            "summary": canonical(asdict(summary)),
        }

    source_dir = source_root / "core" / "src" / "base_audit" / "systems" / "s2_report_collection"
    source_files = (
        "service.py",
        "period_engine.py",
        "external_engine.py",
        "rule_engine.py",
        "validation.py",
        "exporter.py",
        "models.py",
        "importer.py",
    )
    report = {
        "status": "python_baseline_generated",
        "source": {
            "project": str(source_root),
            "module": "PeriodAuditV2Service in-memory component order",
            "sha256": {
                name: hashlib.sha256((source_dir / name).read_bytes()).hexdigest().upper()
                for name in source_files
            },
            "baseline_python": sys.version.split()[0],
        },
        "fixture": {"path": "tests/fixtures/s2_analysis/cases.json", "case_count": len(fixture["cases"])},
        "result": results,
        "scope": (
            "In-memory S2 analysis orchestration and aggregation order only."
            " Mirrors the legacy service component order without calling its filesystem-bound"
            " config check/load, workbook imports, progress callbacks, or XLSX exporter."
        ),
    }
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(
        json.dumps(report, ensure_ascii=False, indent=2, allow_nan=False) + "\n",
        encoding="utf-8",
    )
    print(json.dumps({"output": str(args.output), "case_count": len(results), "source_sha256": report["source"]["sha256"]}, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
