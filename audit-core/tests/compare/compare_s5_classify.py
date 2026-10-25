#!/usr/bin/env python3
"""Compare the Rust S5 business-identification port with the Python baseline.

The comparator collects the real S5 sample subtree files, copies the template
directory and audit configuration into a temporary directory (the source tree
stays read-only; Python's TemplateCatalog writes its index next to the copied
template directory), runs `classify_s5_files` on both sides, and diffs every
row field. Only aggregate counts are persisted.
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

ROOT = Path(__file__).resolve().parents[3]

ROW_FIELDS = (
    "name",
    "template",
    "match_status",
    "business",
    "period",
    "match_detail",
    "matched_sheets",
    "source_sheet_count",
    "template_sheet_count",
)


def collect_files(sample_dir: Path) -> list[Path]:
    files = [
        path
        for path in sorted(sample_dir.rglob("*"))
        if path.is_file()
        and path.suffix.lower() in {".xlsx", ".xlsm"}
        and not path.name.startswith("~$")
    ]
    return files


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--source-root", type=Path, required=True)
    parser.add_argument("--sample-dir", type=Path, required=True,
                        help="real S5 sample subtree (read-only)")
    parser.add_argument("--template-dir", type=Path, required=True,
                        help="S5 template directory inside the sample tree (read-only)")
    parser.add_argument("--config-workbook", type=Path, required=True,
                        help="real S5 audit configuration workbook (read-only)")
    parser.add_argument("--workspace", type=Path, default=ROOT)
    parser.add_argument("--report", type=Path, required=True)
    args = parser.parse_args()

    source_root = args.source_root.resolve(strict=True)
    workspace = args.workspace.resolve(strict=True)
    report_path = args.report.resolve()
    sys.path.insert(0, str(source_root / "core" / "src"))
    from s1audit_configured.workbook_config import load_configuration
    from base_audit.systems.s5_configured_statistics.recognition import (
        classify_s5_files,
    )

    files = collect_files(args.sample_dir.resolve(strict=True))
    if not files:
        raise SystemExit("sample subtree has no Excel files")

    with tempfile.TemporaryDirectory(prefix="audit-s5-classify-") as temporary:
        directory = Path(temporary)
        template_copy = directory / "templates"
        shutil.copytree(args.template_dir.resolve(strict=True), template_copy)
        config_copy = directory / "config.xlsx"
        shutil.copy2(args.config_workbook.resolve(strict=True), config_copy)

        file_strings = [str(path) for path in files]
        python_rows = classify_s5_files(
            file_strings,
            template_dir=str(template_copy),
            config_path=str(config_copy),
        )

        request = {
            "files": file_strings,
            "template_dir": str(template_copy),
            "config_path": str(config_copy),
        }
        request_path = directory / "request.json"
        request_path.write_text(
            json.dumps(request, ensure_ascii=False), encoding="utf-8"
        )
        rust_report = directory / "rust_rows.json"
        environment = os.environ.copy()
        environment["AUDIT_S5_CLASSIFY_INPUT"] = str(request_path)
        environment["AUDIT_S5_CLASSIFY_REPORT"] = str(rust_report)
        completed = subprocess.run(
            [
                "cargo", "test", "-p", "audit-core", "--test", "s5_classify_parity",
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
        rust_rows = json.loads(rust_report.read_text(encoding="utf-8"))

        field_diff_counts: dict[str, int] = {}
        mismatch_ordinals: list[int] = []
        status_distribution: dict[str, int] = {}
        business_distribution: dict[str, int] = {}
        for index, python_row in enumerate(python_rows):
            status_distribution[python_row["match_status"]] = (
                status_distribution.get(python_row["match_status"], 0) + 1
            )
            business_distribution[python_row["business"] or "（空）"] = (
                business_distribution.get(python_row["business"] or "（空）", 0) + 1
            )
            if index >= len(rust_rows):
                field_diff_counts["row_missing_in_rust"] = (
                    field_diff_counts.get("row_missing_in_rust", 0) + 1
                )
                mismatch_ordinals.append(index)
                continue
            rust_row = rust_rows[index]
            row_diffs = 0
            for field in ROW_FIELDS:
                if python_row[field] != rust_row.get(field):
                    field_diff_counts[field] = field_diff_counts.get(field, 0) + 1
                    row_diffs += 1
            if row_diffs:
                mismatch_ordinals.append(index)

        report = {
            "schema_version": 1,
            "suite": "audit.s5_classify_parity",
            "status": "completed" if not mismatch_ordinals else "mismatch",
            "sample_scope": {
                "file_count": len(files),
                "template_dir_copied_to_temporary": True,
                "config_copied_to_temporary": True,
                "template_dir_source": "sample subtree 模板文件_命名区域",
                "selected_files": 0,
                "business_filter": "",
                "period_hint": "",
            },
            "comparison": {
                "python_rows": len(python_rows),
                "rust_rows": len(rust_rows),
                "mismatch_row_count": len(mismatch_ordinals),
                "mismatch_ordinal_sample": mismatch_ordinals[:10],
                "field_diff_counts": dict(sorted(field_diff_counts.items())),
                "compared_fields": list(ROW_FIELDS),
                "python_match_status_distribution": dict(sorted(status_distribution.items())),
                "python_business_distribution": dict(sorted(business_distribution.items())),
            },
            "verification": {
                "rust_runner": "audit-core/tests/s5_classify_parity.rs",
                "python_baseline": (
                    "systems.s5_configured_statistics.recognition.classify_s5_files"
                ),
                "inputs_unchanged": True,
                "raw_paths_or_sheet_names_persisted": False,
                "capability_status": "unverified",
                "production_route_eligible": False,
            },
            "limitations": [
                "TemplateCatalog index persistence is not ported; the Python "
                "baseline writes its index beside the temporary template copy.",
                "match_s5_business propagates workbook XML parse errors in "
                "Python; the Rust port maps read failures to an empty sheet set.",
            ],
        }
        report_path.parent.mkdir(parents=True, exist_ok=True)
        report_path.write_text(
            json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
        )
        print(json.dumps({
            "status": report["status"],
            "rows": len(python_rows),
            "mismatches": len(mismatch_ordinals),
            "field_diffs": dict(sorted(field_diff_counts.items())),
        }))
        return 0 if not mismatch_ordinals else 1


if __name__ == "__main__":
    raise SystemExit(main())
