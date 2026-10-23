#!/usr/bin/env python3
"""Generate a read-only Python baseline for the S2 config loader."""

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
FIXTURE_DIR = WORKSPACE / "tests" / "fixtures" / "s2_config"
FIXTURE_INDEX = FIXTURE_DIR / "cases.json"
DEFAULT_OUTPUT = WORKSPACE / "probe" / "reports" / "s2_config_python_baseline.json"


def canonical(value: object) -> object:
    if isinstance(value, dict):
        return {str(key): canonical(value[key]) for key in sorted(value, key=str)}
    if isinstance(value, tuple):
        return [canonical(item) for item in value]
    if isinstance(value, list):
        return [canonical(item) for item in value]
    return value


def normalize_sources(value: object) -> object:
    if isinstance(value, dict):
        result = {key: normalize_sources(item) for key, item in value.items()}
        if isinstance(result.get("source_file"), str):
            result["source_file"] = result["source_file"].replace("\\", "/").rsplit("/", 1)[-1]
        return result
    if isinstance(value, list):
        return [normalize_sources(item) for item in value]
    return value


def normalized_error(error: Exception, path: Path) -> str:
    message = str(error)
    resolved = str(path.resolve())
    return message.replace(resolved, path.name).replace(resolved.replace("\\", "/"), path.name)


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--source-root", type=Path, default=DEFAULT_SOURCE)
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    args = parser.parse_args()

    source_root = args.source_root.resolve()
    sys.path.insert(0, str(source_root / "core" / "src"))
    from base_audit.systems.s2_report_collection.config import load_config, load_unit_settings

    cases = json.loads(FIXTURE_INDEX.read_text(encoding="utf-8"))["cases"]
    results = {}
    for case in cases:
        path = FIXTURE_DIR / case["file"]
        result: dict[str, object] = {}
        try:
            result["config"] = normalize_sources(canonical(asdict(load_config(path))))
        except Exception as error:  # the error text is part of this comparison contract
            result["config_error"] = normalized_error(error, path)
        try:
            result["unit_settings"] = canonical(asdict(load_unit_settings(path)))
        except Exception as error:  # unit-only loader errors are compared separately
            result["unit_settings_error"] = normalized_error(error, path)
        results[case["id"]] = result

    config_module = source_root / "core" / "src" / "base_audit" / "systems" / "s2_report_collection" / "config.py"
    models_module = source_root / "core" / "src" / "base_audit" / "systems" / "s2_report_collection" / "models.py"
    units_module = source_root / "core" / "src" / "base_audit" / "systems" / "s2_report_collection" / "unit_conversion.py"
    source_hashes = {
        path.name: hashlib.sha256(path.read_bytes()).hexdigest().upper()
        for path in (config_module, models_module, units_module)
    }
    report = {
        "status": "python_baseline_generated",
        "source": {
            "project": str(source_root),
            "module": "base_audit.systems.s2_report_collection.config.{load_config,load_unit_settings}",
            "source_sha256": source_hashes,
            "baseline_python": sys.version.split()[0],
        },
        "fixture": {"path": "tests/fixtures/s2_config/cases.json", "case_count": len(cases)},
        "result": results,
        "scope": "S2 runtime config workbook parsing and shared unit-settings loading over synthetic XLSX fixtures. Does not cover config validation UI/checker behavior, default-config XLSX generation, write-back, service orchestration, or real business configuration files.",
    }
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(report, ensure_ascii=False, indent=2, allow_nan=False) + "\n", encoding="utf-8")
    print(json.dumps({
        "output": str(args.output),
        "case_count": len(cases),
        "config_success_count": sum("config" in value for value in results.values()),
        "config_error_count": sum("config_error" in value for value in results.values()),
        "unit_settings_success_count": sum("unit_settings" in value for value in results.values()),
        "source_sha256": source_hashes,
    }, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
