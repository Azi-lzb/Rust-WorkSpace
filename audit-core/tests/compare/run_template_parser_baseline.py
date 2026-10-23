"""Generate a read-only Python baseline for base_audit.template pure helpers."""

from __future__ import annotations

import argparse
import hashlib
import json
import sys
from dataclasses import asdict
from pathlib import Path


WORKSPACE = Path(__file__).resolve().parents[2]
CASES = WORKSPACE / "tests/fixtures/template_parser_cases.json"
SOURCE_RELATIVE = Path("core/src/base_audit/template.py")


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--source-root",
        type=Path,
        default=Path(r"C:\Users\AZI\Desktop\VibeCode\3.open-data-audit"),
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=WORKSPACE / "probe/reports/template_parser_python_baseline.json",
    )
    args = parser.parse_args()
    sys.dont_write_bytecode = True
    sys.path.insert(0, str(args.source_root / "core/src"))

    from base_audit.models import AuditRule
    from base_audit.template import (
        TemplateError,
        clean_rule_comment,
        normalize_template_name,
        parse_comment_rule_id,
        parse_formula_result,
        parse_rule_rows,
    )

    cases = json.loads(CASES.read_text(encoding="utf-8"))
    comment_cases = []
    for case in cases["comment_cases"]:
        try:
            rule_id = {"value": parse_comment_rule_id(case["text"])}
        except TemplateError as error:
            rule_id = {"error": str(error)}
        comment_cases.append(
            {
                "name": case["name"],
                "rule_id": rule_id,
                "cleaned": clean_rule_comment(case["text"]),
            }
        )

    formula_cases = []
    for case in cases["formula_cases"]:
        parsed = parse_formula_result(
            case.get("value"),
            default_severity=case.get("default_severity", "错误"),
            default_message=case.get("default_message", ""),
        )
        formula_cases.append({"name": case["name"], "result": asdict(parsed)})

    rule_cases = []
    for case in cases["rule_row_cases"]:
        try:
            result = {"rules": [asdict(rule) for rule in parse_rule_rows(case["rows"])]}
        except TemplateError as error:
            result = {"error": str(error)}
        rule_cases.append({"name": case["name"], "result": result})

    source_file = args.source_root / SOURCE_RELATIVE
    payload = {
        "source": {
            "project": str(args.source_root),
            "module": SOURCE_RELATIVE.as_posix(),
            "sha256": hashlib.sha256(source_file.read_bytes()).hexdigest().upper(),
        },
        "normalize_names": [
            {"input": value, "output": normalize_template_name(value)}
            for value in cases["normalize_names"]
        ],
        "comment_cases": comment_cases,
        "formula_cases": formula_cases,
        "rule_row_cases": rule_cases,
    }
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )
    print(args.output)


if __name__ == "__main__":
    main()
