#!/usr/bin/env python3
"""Read-only Python baseline for the Audit conditional-rule evaluator."""

from __future__ import annotations

import argparse
import hashlib
import json
import sys
from pathlib import Path


SOURCE_CORE = Path(r"C:\Users\AZI\Desktop\VibeCode\3.open-data-audit\core")
sys.path.insert(0, str(SOURCE_CORE / "src"))

from base_audit.cf_reader import OoxmlConditionalRuleReader  # noqa: E402
from base_audit.conditional_evaluators import (  # noqa: E402
    PythonConditionalRuleEvaluator,
)


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--workbook", required=True, type=Path)
    parser.add_argument("--output", type=Path)
    args = parser.parse_args()
    workbook = args.workbook.resolve(strict=True)

    reader = OoxmlConditionalRuleReader()
    reader_result = reader.read(workbook)
    evaluator = PythonConditionalRuleEvaluator()
    rules_by_sheet: dict[str, list] = {}
    for rule in reader_result.rules:
        rules_by_sheet.setdefault(rule.sheet, []).append(rule)

    triggers = []
    unsupported_cells = 0
    error_count = 0
    evaluated_cells = 0
    for sheet, rules in rules_by_sheet.items():
        bounds = list(dict.fromkeys(bound for rule in rules for bound in rule.bounds))
        results, stats = evaluator.evaluate_sheet(
            sheet, rules, bounds, reader_result
        )
        evaluated_cells += stats["evaluated_cells"]
        unsupported_cells += len(stats["unsupported_cells"])
        error_count += len(stats["errors"])
        triggers.extend(
            {
                "sheet": result.sheet,
                "cell": result.cell,
                "applies_to": result.range,
                "rule_type": result.rule_type,
                "formula": result.formula,
                "triggered": result.triggered,
            }
            for result in results
        )

    report = {
        "schema_version": 1,
        "baseline": "Open-Data-Audit PythonConditionalRuleEvaluator",
        "workbook_sha256": sha256(workbook),
        "source_sha256": {
            "cf_reader.py": sha256(SOURCE_CORE / "src/base_audit/cf_reader.py"),
            "conditional_evaluators.py": sha256(
                SOURCE_CORE / "src/base_audit/conditional_evaluators.py"
            ),
            "conditional_format.py": sha256(
                SOURCE_CORE / "src/base_audit/conditional_format.py"
            ),
        },
        "rules_total": len(reader_result.rules),
        "rules_supported": sum(rule.supported for rule in reader_result.rules),
        "evaluated_cells": evaluated_cells,
        "triggered": sorted(triggers, key=lambda row: (row["sheet"], row["cell"])),
        "unsupported_cells": unsupported_cells,
        "error_count": error_count,
        "diagnostics_by_status": {
            "unsupported_cells": unsupported_cells,
            "errors": error_count,
        },
    }
    encoded = json.dumps(report, ensure_ascii=False, indent=2) + "\n"
    if args.output:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(encoded, encoding="utf-8")
    else:
        sys.stdout.write(encoded)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
