from __future__ import annotations

import argparse
import importlib
import json
import logging
import subprocess
import sys
import tempfile
from pathlib import Path

from compare_sheet_extract_area_durations import create_rust_config, project_workbook


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--pytools-cli", type=Path, required=True)
    parser.add_argument("--python-project", type=Path, required=True)
    parser.add_argument("--fixture", type=Path)
    parser.add_argument("--report", type=Path)
    args = parser.parse_args()

    probe_root = Path(__file__).resolve().parent
    pytools_root = probe_root.parent
    fixture = args.fixture or (
        pytools_root
        / "tests"
        / "fixtures"
        / "sheet_extract_area"
        / "iso-typed-durations.xlsx"
    )
    report_path = args.report or (
        probe_root / "reports" / "sheet_extract_iso_duration_parity_20261026.json"
    )
    source_python = args.python_project.resolve() / "PyCode"
    sys.path.insert(0, str(source_python))
    legacy = importlib.import_module("core.feature_1_6_extract_sheets")
    source_bytes = fixture.read_bytes()

    with tempfile.TemporaryDirectory(prefix="pytools-iso-duration-parity-") as temporary:
        root = Path(temporary)
        python_output = root / "python-output"
        python_logs = root / "python-logs"
        rust_output = root / "rust-output"
        config_path = root / "config.xlsx"
        row_count = 5
        task = {
            "row_no": 2,
            "workbook_keywords": "",
            "sheet_keywords": "",
            "exclude_keywords": "",
            "area_spec": f"A1:A{row_count}",
            "stitch_mode": "",
            "output_folder": "时长提取",
            "output_name": "duration-output.xlsx",
            "remark": "strict OOXML ISO duration parity",
        }

        try:
            python_result = legacy.run_extract_sheets(
                [fixture],
                [task],
                python_output,
                python_logs,
                append_existing=False,
                default_output_name="unused.xlsx",
            )
            create_rust_config(config_path, rust_output, row_count)
            rust_process = subprocess.run(
                [
                    str(args.pytools_cli.resolve()),
                    "extract-sheets-area",
                    "--config",
                    str(config_path),
                    "--output-dir",
                    str(rust_output),
                    str(fixture),
                ],
                capture_output=True,
                text=True,
                check=False,
            )
            if rust_process.returncode != 0:
                raise RuntimeError(
                    f"Rust CLI failed ({rust_process.returncode}): {rust_process.stderr}"
                )
            rust_stats = json.loads(rust_process.stdout)
            python_projection = project_workbook(
                python_output / "时长提取" / "duration-output.xlsx"
            )
            rust_projection = project_workbook(
                rust_output / "时长提取" / "duration-output.xlsx"
            )
            source_unchanged = fixture.read_bytes() == source_bytes
            if not source_unchanged:
                raise AssertionError("synthetic ISO duration fixture changed")
            if python_projection != rust_projection:
                raise AssertionError(
                    "Python/Rust output mismatch:\n"
                    + json.dumps(
                        {"python": python_projection, "rust": rust_projection},
                        ensure_ascii=False,
                        indent=2,
                    )
                )
            if python_result["failed_files"] != 0 or rust_stats["failed_files"] != 0:
                raise AssertionError("the ISO duration extraction task reported a failed source")

            report = {
                "schema_version": 1,
                "status": "passed",
                "feature": "PT-7-6 area extraction strict OOXML ISO duration subset",
                "sample_kind": "synthetic XLSX cells tagged t=d with PT durations",
                "source_python": "PyCode/core/feature_1_6_extract_sheets.py::run_extract_sheets",
                "rust_entry": "pytools extract-sheets-area",
                "python_stats": {
                    key: value for key, value in python_result.items() if key != "paths"
                },
                "rust_stats": {
                    key: value for key, value in rust_stats.items() if key != "paths"
                },
                "output_projection": python_projection,
                "source_unchanged": source_unchanged,
                "scope_note": (
                    "Synthetic parity only. Real business files, malformed or unsupported ISO "
                    "duration forms, time-only cells, and platform delivery remain unverified."
                ),
            }
            report_path.parent.mkdir(parents=True, exist_ok=True)
            report_path.write_text(
                json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
            )
            print(json.dumps(report, ensure_ascii=False, indent=2))
        finally:
            logging.shutdown()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
