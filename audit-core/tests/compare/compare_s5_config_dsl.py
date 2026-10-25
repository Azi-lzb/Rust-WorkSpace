#!/usr/bin/env python3
"""Compare the Rust S5 configuration DSL port with the Python baseline.

The comparator extracts the DSL fields of every enabled rule from the real S5
audit configuration workbook, replays the source project's `load_configuration`
DSL check sequence per rule, runs the Rust parity runner over the same items,
and compares status, error class, and raw message field by field. The report
persists aggregate counts only; no rule text, expression, or error message is
written.
"""

from __future__ import annotations

import argparse
import json
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
ROOT = Path(__file__).resolve().parents[3]


def _text(value) -> str:
    if value is None:
        return ""
    return str(value).strip()


def _flag_text(value) -> bool:
    text = _text(value).lower()
    if text in {"是", "是的", "yes", "y", "true", "1", "启用", "已确认"}:
        return True
    if text in {"否", "不是", "no", "n", "false", "0", "禁用", "未确认"}:
        return False
    raise ValueError("是否启用 必须填写 是/否")


EXPRESSION_FIELDS = (
    "审核表达式",
    "当前值表达式",
    "对比值表达式",
    "差值表达式",
    "适用表达式",
)


def extract_enabled_rules(config_workbook: Path):
    from openpyxl import load_workbook
    from openpyxl.utils import get_column_letter

    workbook = load_workbook(config_workbook, data_only=False, read_only=False)
    sheet = workbook["审核规则"]
    header_values = []
    for column in range(1, sheet.max_column + 1):
        header_values.append(_text(sheet.cell(1, column).value))
    while header_values and not header_values[-1]:
        header_values.pop()
    columns = {header: index + 1 for index, header in enumerate(header_values)}

    rows = []
    formula_rows = 0
    for row_num in range(2, sheet.max_row + 1):
        values = {
            header: sheet.cell(row_num, columns[header]).value for header in columns
        }
        if not any(value is not None and _text(value) for value in values.values()):
            continue
        for header, column in columns.items():
            if sheet.cell(row_num, column).data_type == "f":
                formula_rows += 1
                if formula_rows == 1:
                    print(
                        json.dumps(
                            {"config_note": "formula_cell_rejected", "row": row_num,
                             "column": get_column_letter(column)}
                        )
                    )
        rows.append((row_num, values))
    workbook.close()
    return rows, formula_rows


def python_check_rule(values) -> dict:
    """Replay the DSL segment of the source `load_configuration` rule loop."""
    from s1audit_configured.dsl import (
        parse_dictionary,
        parse_expression,
        static_kind,
        validate_dictionary_dependencies,
        validate_static_types,
        variables_in,
    )
    from s1audit_configured.errors import ConfigurationError, DSLParseError

    def outcome(error):
        return {
            "status": "error",
            "error_kind": "parse" if isinstance(error, DSLParseError) else "configuration",
            "message": error.message,
        }

    try:
        dictionary_text = _text(values.get("取值字典")) or None
        dictionary_nodes = parse_dictionary(dictionary_text)
        validate_dictionary_dependencies(dictionary_nodes)
        for node in dictionary_nodes.values():
            validate_static_types(node)
        expression_nodes = {}
        for field_name in EXPRESSION_FIELDS:
            expression = _text(values.get(field_name))
            if not expression and field_name != "审核表达式":
                continue
            expression_nodes[field_name] = parse_expression(expression)
        for field_name, node in expression_nodes.items():
            validate_static_types(node)
            kind = static_kind(node)
            if field_name in {"审核表达式", "适用表达式"} and kind not in {None, "boolean"}:
                return outcome(ConfigurationError(f"{field_name} 必须返回布尔值"))
            if field_name == "差值表达式" and kind not in {None, "number"}:
                return outcome(ConfigurationError("差值表达式必须返回数值"))
        for field_name, node in expression_nodes.items():
            undefined = variables_in(node) - dictionary_nodes.keys()
            if undefined:
                return outcome(ConfigurationError(
                    f"{field_name} 引用未定义变量：" + ", ".join(sorted(undefined))
                ))
        return {"status": "ok", "error_kind": None, "message": None}
    except ConfigurationError as error:
        return outcome(error)


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--source-root", type=Path, required=True)
    parser.add_argument("--config-workbook", type=Path, required=True,
                        help="real S5 audit configuration workbook (.xlsx or byte-copy)")
    parser.add_argument("--workspace", type=Path, default=ROOT)
    parser.add_argument("--report", type=Path, required=True)
    args = parser.parse_args()

    source_root = args.source_root.resolve(strict=True)
    workspace = args.workspace.resolve(strict=True)
    report_path = args.report.resolve()

    sys.path.insert(0, str(source_root / "core" / "src"))
    from s1audit_configured.dsl import (  # noqa: F401  (import sanity)
        parse_expression,
    )

    with tempfile.TemporaryDirectory(prefix="audit-s5-dsl-parity-") as temporary:
        directory = Path(temporary)
        config_copy = directory / "config.xlsx"
        source = args.config_workbook.resolve(strict=True)
        shutil.copy2(source, config_copy)

        rows, formula_rows = extract_enabled_rules(config_copy)

        items = []
        python_outcomes = {}
        enabled = 0
        missing_expression = 0
        for row_num, values in rows:
            if not _flag_text(values.get("是否启用")):
                continue
            enabled += 1
            audit_expression = _text(values.get("审核表达式"))
            if not audit_expression:
                missing_expression += 1
                continue
            expressions = [("审核表达式", audit_expression)]
            for field_name in EXPRESSION_FIELDS[1:]:
                text = _text(values.get(field_name))
                if text:
                    expressions.append((field_name, text))
            item_id = len(items) + 1
            items.append({
                "id": item_id,
                "dictionary": _text(values.get("取值字典")) or None,
                "expressions": expressions,
            })
            python_outcomes[item_id] = python_check_rule(values)

        input_path = directory / "items.jsonl"
        input_path.write_text(
            "".join(json.dumps(item, ensure_ascii=False) + "\n" for item in items),
            encoding="utf-8",
        )
        rust_report = directory / "rust_outcomes.jsonl"
        environment = os.environ.copy()
        environment["AUDIT_S5_DSL_PARITY_INPUT"] = str(input_path)
        environment["AUDIT_S5_DSL_PARITY_REPORT"] = str(rust_report)
        completed = subprocess.run(
            [
                "cargo", "test", "-p", "audit-core", "--test", "s5_config_dsl_parity",
                "--locked", "--", "--ignored", "--nocapture",
            ],
            cwd=workspace,
            env=environment,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            check=False,
        )
        if completed.returncode != 0 or not rust_report.is_file():
            print(json.dumps({
                "status": "rust_runner_failed",
                "exit_code": completed.returncode,
                "stderr_tail": completed.stderr[-2000:],
            }))
            return 1

        rust_outcomes = {
            json.loads(line)["id"]: json.loads(line)
            for line in rust_report.read_text(encoding="utf-8").splitlines()
            if line.strip()
        }

        matched = 0
        status_mismatch_ids = []
        message_mismatch_ids = []
        kind_distribution = {}
        for item in items:
            identifier = item["id"]
            expected = python_outcomes[identifier]
            actual = rust_outcomes.get(identifier)
            if actual is None:
                status_mismatch_ids.append(identifier)
                continue
            status_match = expected["status"] == actual["status"]
            message_match = expected["message"] == actual["message"]
            if status_match and message_match and expected["error_kind"] == actual["error_kind"]:
                matched += 1
            else:
                if not status_match:
                    status_mismatch_ids.append(identifier)
                elif not message_match:
                    message_mismatch_ids.append(identifier)
            key = f"python_{expected['error_kind'] or 'ok'}/rust_{(actual or {}).get('error_kind') or 'ok'}"
            kind_distribution[key] = kind_distribution.get(key, 0) + 1

        mismatch_count = len(status_mismatch_ids) + len(message_mismatch_ids)
        report = {
            "schema_version": 1,
            "suite": "audit.s5_config_dsl_parity",
            "status": "completed" if mismatch_count == 0 else "mismatch",
            "sample_scope": {
                "rule_rows_total": len(rows),
                "enabled_rules": enabled,
                "compared_rules": len(items),
                "missing_audit_expression_skipped": missing_expression,
                "formula_cells_rejected_rows": formula_rows,
            },
            "comparison": {
                "matched": matched,
                "status_mismatch_count": len(status_mismatch_ids),
                "message_mismatch_count": len(message_mismatch_ids),
                "mismatch_ordinal_sample": sorted(
                    set(status_mismatch_ids) | set(message_mismatch_ids)
                )[:10],
                "outcome_kind_distribution": dict(sorted(kind_distribution.items())),
                "compared_fields": ["status", "error_kind", "message"],
            },
            "verification": {
                "rust_runner": "audit-core/tests/s5_config_dsl_parity.rs",
                "python_baseline": "s1audit_configured.dsl via load_configuration rule-loop semantics",
                "inputs_unchanged": True,
                "raw_rule_text_or_messages_persisted": False,
                "capability_status": "unverified",
                "production_route_eligible": False,
            },
            "limitations": [
                "Covers the DSL parse/validate layer only; workbook loading, "
                "scope/business row checks, and rule evaluation are not migrated.",
                "Python str.isdigit accepts all Unicode decimal digits; the Rust "
                "port accepts ASCII plus the superscript trio (documented delta).",
            ],
        }
        report_path.parent.mkdir(parents=True, exist_ok=True)
        report_path.write_text(
            json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
        )
        print(json.dumps({
            "status": report["status"],
            "compared": len(items),
            "matched": matched,
            "mismatches": mismatch_count,
        }))
        return 0 if mismatch_count == 0 and formula_rows == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())
