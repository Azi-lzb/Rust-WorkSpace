#!/usr/bin/env python3
"""Compare the Rust S5 configuration loader port with the Python baseline.

Both sides load the same byte-copy of the real audit configuration workbook.
The Python baseline runs `s1audit_configured.workbook_config.load_configuration`
directly; the Rust side runs the `audit_core::s5_config` port through the
manual parity test. The comparator diffs the normalized `Configuration`
structures field by field and writes aggregate counts only.
"""

from __future__ import annotations

import argparse
import dataclasses
import json
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[3]


def normalize_configuration(configuration) -> dict:
    payload = dataclasses.asdict(configuration)
    payload.pop("path", None)
    # Round-trip through JSON so tuples become lists and match the Rust side.
    return json.loads(json.dumps(payload, ensure_ascii=False))


def diff_values(left, right, path: str, diffs: list[str]) -> None:
    if isinstance(left, dict) and isinstance(right, dict):
        for key in sorted(set(left) | set(right)):
            if key not in left:
                diffs.append(f"{path}.{key}: missing_in_python")
            elif key not in right:
                diffs.append(f"{path}.{key}: missing_in_rust")
            else:
                diff_values(left[key], right[key], f"{path}.{key}", diffs)
    elif isinstance(left, list) and isinstance(right, list):
        if len(left) != len(right):
            diffs.append(f"{path}: length {len(left)} vs {len(right)}")
        for index in range(min(len(left), len(right))):
            diff_values(left[index], right[index], f"{path}[{index}]", diffs)
    elif left != right:
        diffs.append(path)


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--source-root", type=Path, required=True)
    parser.add_argument("--config-workbook", type=Path, required=True)
    parser.add_argument("--workspace", type=Path, default=ROOT)
    parser.add_argument("--report", type=Path, required=True)
    args = parser.parse_args()

    source_root = args.source_root.resolve(strict=True)
    workspace = args.workspace.resolve(strict=True)
    report_path = args.report.resolve()
    sys.path.insert(0, str(source_root / "core" / "src"))
    from s1audit_configured.workbook_config import load_configuration

    with tempfile.TemporaryDirectory(prefix="audit-s5-config-load-") as temporary:
        directory = Path(temporary)
        config_copy = directory / "config.xlsx"
        shutil.copy2(args.config_workbook.resolve(strict=True), config_copy)

        try:
            configuration = load_configuration(config_copy, include_disabled=False, business="")
            python_payload = {
                "status": "ok",
                "configuration": normalize_configuration(configuration),
            }
        except Exception as error:  # ConfigurationError carries .message
            message = getattr(error, "message", str(error))
            python_payload = {"status": "error", "message": message}

        rust_report = directory / "rust_configuration.json"
        environment = os.environ.copy()
        environment["AUDIT_S5_CONFIG_LOAD_INPUT"] = str(config_copy)
        environment["AUDIT_S5_CONFIG_LOAD_REPORT"] = str(rust_report)
        completed = subprocess.run(
            [
                "cargo", "test", "-p", "audit-core", "--test", "s5_config_load_parity",
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
        rust_payload = json.loads(rust_report.read_text(encoding="utf-8"))

        diffs: list[str] = []
        if python_payload["status"] != rust_payload.get("status"):
            diffs.append("status")
        else:
            if python_payload["status"] == "error":
                if python_payload.get("message") != rust_payload.get("message"):
                    diffs.append("message")
            else:
                diff_values(
                    python_payload["configuration"],
                    rust_payload.get("configuration", {}),
                    "configuration",
                    diffs,
                )

        field_counts: dict[str, int] = {}
        for diff in diffs:
            head = diff.split(":")[0].split("[")[0]
            field_counts[head] = field_counts.get(head, 0) + 1

        rules = python_payload.get("configuration", {}).get("rules", [])
        report = {
            "schema_version": 1,
            "suite": "audit.s5_config_load_parity",
            "status": "completed" if not diffs else "mismatch",
            "sample_scope": {
                "config_sheet_shape": "rules sheet without optional catalog sheets",
                "python_rule_count": len(rules),
                "include_disabled": False,
                "business_filter": "",
            },
            "comparison": {
                "diff_count": len(diffs),
                "field_diff_counts": dict(sorted(field_counts.items())),
                "diff_path_sample": sorted(set(diffs))[:20],
                "compared_fields": [
                    "rules(all fields)",
                    "metrics",
                    "external_metrics",
                    "source_contracts",
                    "applicability_scopes",
                    "version",
                    "version_status",
                    "effective_from",
                    "effective_to",
                    "release_date",
                    "release_note",
                    "sha256",
                ],
            },
            "verification": {
                "rust_runner": "audit-core/tests/s5_config_load_parity.rs",
                "python_baseline": "s1audit_configured.workbook_config.load_configuration",
                "inputs_unchanged": True,
                "raw_rule_text_or_messages_persisted": False,
                "capability_status": "unverified",
                "production_route_eligible": False,
            },
            "limitations": [
                "Optional catalog sheets (模板指标目录/适用范围目录/外部指标目录/数据源契约/"
                "配置版本) are absent from the real configuration, so only their "
                "absence branches plus derived metric/external extraction are "
                "exercised by this parity run.",
                "The workbook-not-found and workbook-unreadable messages embed "
                "runtime paths and are not compared verbatim.",
            ],
        }
        report_path.parent.mkdir(parents=True, exist_ok=True)
        report_path.write_text(
            json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
        )
        print(json.dumps({
            "status": report["status"],
            "python_status": python_payload["status"],
            "diff_count": len(diffs),
        }))
        return 0 if not diffs else 1


if __name__ == "__main__":
    raise SystemExit(main())
