from __future__ import annotations

import argparse
import importlib
import json
import logging
import subprocess
import sys
import tempfile
from datetime import time
from pathlib import Path

from openpyxl import Workbook

from compare_sheet_extract_area_durations import create_rust_config, project_workbook


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--pytools-cli", type=Path, required=True)
    parser.add_argument("--python-project", type=Path, required=True)
    parser.add_argument("--report", type=Path)
    args = parser.parse_args()

    sys.dont_write_bytecode = True
    probe_root = Path(__file__).resolve().parent
    report_path = args.report or (
        probe_root / "reports" / "sheet_extract_area_time_only_blocker_20261026.json"
    )
    source_python = args.python_project.resolve() / "PyCode"
    sys.path.insert(0, str(source_python))
    legacy = importlib.import_module("core.feature_1_6_extract_sheets")

    with tempfile.TemporaryDirectory(prefix="pytools-time-only-blocker-") as temporary:
        root = Path(temporary)
        source = root / "time-only.xlsx"
        workbook = Workbook()
        worksheet = workbook.active
        worksheet["A1"] = time(0, 0, 0)
        worksheet["A1"].number_format = "h:mm:ss"
        workbook.save(source)
        workbook.close()
        source_bytes = source.read_bytes()

        python_output = root / "python-output"
        python_logs = root / "python-logs"
        rust_output = root / "rust-output"
        config_path = root / "config.xlsx"
        task = {
            "row_no": 2,
            "workbook_keywords": "",
            "sheet_keywords": "",
            "exclude_keywords": "",
            "area_spec": "A1",
            "stitch_mode": "",
            "output_folder": "时分秒",
            "output_name": "time-output.xlsx",
            "remark": "strict OOXML time-only reproducer",
        }

        try:
            python_result = legacy.run_extract_sheets(
                [source],
                [task],
                python_output,
                python_logs,
                append_existing=False,
                default_output_name="unused.xlsx",
            )
            create_rust_config(config_path, rust_output, 1)
            rust_process = subprocess.run(
                [
                    str(args.pytools_cli.resolve()),
                    "extract-sheets-area",
                    "--config",
                    str(config_path),
                    "--output-dir",
                    str(rust_output),
                    str(source),
                ],
                capture_output=True,
                text=True,
                check=False,
            )

            python_path = python_output / "时分秒" / "time-output.xlsx"
            rust_path = rust_output / "时分秒" / "time-output.xlsx"
            python_projection = project_workbook(python_path)
            source_unchanged = source.read_bytes() == source_bytes
            if not source_unchanged:
                raise AssertionError("synthetic time-only source workbook changed")
            if python_result["failed_files"] != 0:
                raise AssertionError("PyCode rejected the synthetic time-only workbook")

            if rust_process.returncode == 0 and rust_path.is_file():
                rust_projection = project_workbook(rust_path)
                status = "passed" if python_projection == rust_projection else "failed"
            else:
                rust_projection = None
                expected_error = "ambiguous with a time-only value"
                status = (
                    "blocked"
                    if rust_process.returncode == 2
                    and expected_error in rust_process.stderr
                    and not rust_path.exists()
                    else "failed"
                )

            report = {
                "schema_version": 1,
                "status": status,
                "feature": "PT-7-6 area extraction time-only value classification",
                "sample_kind": "synthetic XLSX time value with h:mm:ss format",
                "source_python": "PyCode/core/feature_1_6_extract_sheets.py::run_extract_sheets",
                "rust_entry": "pytools extract-sheets-area",
                "shared_api_gap": (
                    "excel-reader returns CellValue without per-cell number-format metadata, "
                    "so PyTools cannot distinguish a time-only value from an epoch date."
                ),
                "python_stats": {
                    key: value for key, value in python_result.items() if key != "paths"
                },
                "python_projection": python_projection,
                "rust_returncode": rust_process.returncode,
                "rust_error": rust_process.stderr.strip(),
                "rust_output_created": rust_path.is_file(),
                "rust_projection": rust_projection,
                "source_unchanged": source_unchanged,
                "scope_note": (
                    "Synthetic reproducer only. Real business time-only cells and other "+
                    "formats/platforms remain unverified. A reader metadata API is needed "+
                    "before implementing value classification without a PyTools-local OOXML reader."
                ),
            }
            report_path.parent.mkdir(parents=True, exist_ok=True)
            report_path.write_text(
                json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
            )
            print(json.dumps(report, ensure_ascii=False, indent=2))
            return 0 if status in {"blocked", "passed"} else 1
        finally:
            logging.shutdown()


if __name__ == "__main__":
    raise SystemExit(main())
