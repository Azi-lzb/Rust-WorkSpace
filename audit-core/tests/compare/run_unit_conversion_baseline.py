#!/usr/bin/env python3
"""Generate a read-only baseline for S2 amount unit conversion helpers."""

from __future__ import annotations

import argparse
import hashlib
import importlib
import importlib.util
import json
import sys
from decimal import Decimal
from pathlib import Path


WORKSPACE = Path(__file__).resolve().parents[2]
CASES = WORKSPACE / "tests/fixtures/unit_conversion_cases.json"
SOURCE_RELATIVE = Path(
    "core/src/base_audit/systems/s2_report_collection/unit_conversion.py"
)
IMPORTER_RELATIVE = Path(
    "core/src/base_audit/systems/s2_report_collection/importer.py"
)


def load_source(source_root: Path):
    module_path = source_root / SOURCE_RELATIVE
    spec = importlib.util.spec_from_file_location("audit_unit_conversion_baseline", module_path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"cannot load source module: {module_path}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


def load_importer(source_root: Path):
    source_package_root = str(source_root / "core" / "src")
    if source_package_root not in sys.path:
        sys.path.insert(0, source_package_root)
    return importlib.import_module(
        "base_audit.systems.s2_report_collection.importer"
    )


def decimal_tuple(value: Decimal) -> dict[str, object]:
    value_tuple = value.as_tuple()
    return {
        "sign": value_tuple.sign,
        "digits": list(value_tuple.digits),
        "exponent": value_tuple.exponent,
    }


def outcome(call) -> dict[str, object]:
    try:
        return {"status": "ok", "tuple": decimal_tuple(call())}
    except ValueError as error:
        return {
            "status": "error",
            "type": type(error).__name__,
            "message": str(error),
        }


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--source-root",
        type=Path,
        default=Path(r"C:\Users\AZI\Desktop\VibeCode\3.open-data-audit"),
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=WORKSPACE / "probe/reports/unit_conversion_python_baseline.json",
    )
    args = parser.parse_args()
    sys.dont_write_bytecode = True
    source_root = args.source_root.resolve()
    module = load_source(source_root)
    importer = load_importer(source_root)
    cases = json.loads(CASES.read_text(encoding="utf-8"))
    importer_path = source_root / IMPORTER_RELATIVE

    reader_cell_cases = []
    for case in cases["reader_cell_cases"]:
        source_format = case["source_format"]
        normalized = importer._normalize_fast_value(
            case["python_value"], xls=source_format == "xls"
        )
        number = importer._decimal_number(normalized)
        if number is None:
            decimal_output = {"status": "skipped"}
            conversion_output = {"status": "skipped"}
        else:
            decimal_output = {"status": "ok", "tuple": decimal_tuple(number)}
            conversion_output = outcome(
                lambda number=number, item=case: importer.convert_amount(
                    number, item["source_unit"], item["target_unit"]
                )
            )
        reader_cell_cases.append(
            {
                **case,
                "decimal_output": decimal_output,
                "conversion_output": conversion_output,
            }
        )

    report = {
        "source": {
            "project": str(source_root),
            "module": SOURCE_RELATIVE.as_posix(),
            "sha256": hashlib.sha256(
                (source_root / SOURCE_RELATIVE).read_bytes()
            ).hexdigest().upper(),
            "baseline_python": sys.version.split()[0],
            "importer_module": IMPORTER_RELATIVE.as_posix(),
            "importer_sha256": hashlib.sha256(importer_path.read_bytes())
            .hexdigest()
            .upper(),
        },
        "amount_data_type_cases": [
            {
                **case,
                "output": case["value"] in module.AMOUNT_DATA_TYPES,
            }
            for case in cases["amount_data_type_cases"]
        ],
        "as_decimal_cases": [
            {
                **case,
                "output": outcome(lambda value=case["value"]: module.as_decimal(value)),
            }
            for case in cases["as_decimal_cases"]
        ],
        "conversion_cases": [
            {
                **case,
                "output": outcome(
                    lambda item=case: module.convert_amount(
                        item["value"], item["source_unit"], item["target_unit"]
                    )
                ),
            }
            for case in cases["conversion_cases"]
        ],
        "to_yuan_cases": [
            {
                **case,
                "output": outcome(
                    lambda item=case: module.to_yuan(item["value"], item["unit"])
                ),
            }
            for case in cases["to_yuan_cases"]
        ],
        "reader_cell_cases": reader_cell_cases,
    }
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(
        json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )
    print(
        json.dumps(
            {
                "baseline": str(args.output),
                "cases": sum(
                    len(report[key])
                    for key in (
                        "amount_data_type_cases",
                        "as_decimal_cases",
                        "conversion_cases",
                        "to_yuan_cases",
                        "reader_cell_cases",
                    )
                ),
                "source_sha256": report["source"]["sha256"],
            },
            ensure_ascii=False,
        )
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
