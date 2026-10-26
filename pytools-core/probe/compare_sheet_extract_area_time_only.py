from __future__ import annotations

import argparse
import importlib
import json
import logging
import subprocess
import sys
import tempfile
from pathlib import Path

from openpyxl import Workbook
from openpyxl.utils.datetime import MAC_EPOCH

from compare_sheet_extract_area_durations import create_rust_config, project_workbook


def create_source(path: Path, date_1904: bool) -> bytes:
    workbook = Workbook()
    if date_1904:
        workbook.epoch = MAC_EPOCH
    worksheet = workbook.active
    cases = [
        (0.0, "h:mm:ss"),
        (0.5 / 86_400_000, "h:mm:ss"),
        (1.5 / 86_400_000, "h:mm:ss"),
        (86_399_999.5 / 86_400_000, "h:mm:ss"),
        (0.5, "yyyy-mm-dd"),
        (0.5, "h:mm"),
        (0.5, "[h]:mm:ss"),
    ]
    for row, (value, number_format) in enumerate(cases, start=1):
        worksheet.cell(row=row, column=1, value=value).number_format = number_format
    workbook.save(path)
    workbook.close()
    return path.read_bytes()


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--pytools-cli", type=Path, required=True)
    parser.add_argument("--python-project", type=Path, required=True)
    parser.add_argument("--report", type=Path)
    args = parser.parse_args()

    sys.dont_write_bytecode = True
    probe_root = Path(__file__).resolve().parent
    report_path = args.report or (
        probe_root / "reports" / "sheet_extract_area_time_only_parity_20261026.json"
    )
    source_python = args.python_project.resolve() / "PyCode"
    sys.path.insert(0, str(source_python))
    legacy = importlib.import_module("core.feature_1_6_extract_sheets")

    with tempfile.TemporaryDirectory(prefix="pytools-time-only-blocker-") as temporary:
        root = Path(temporary)
        sources = [root / "time-only-1900.xlsx", root / "time-only-1904.xlsx"]
        source_snapshots = {
            source: create_source(source, date_1904=index == 1)
            for index, source in enumerate(sources)
        }

        python_output = root / "python-output"
        python_logs = root / "python-logs"
        rust_output = root / "rust-output"
        config_path = root / "config.xlsx"
        task = {
            "row_no": 2,
            "workbook_keywords": "",
            "sheet_keywords": "",
            "exclude_keywords": "",
            "area_spec": "A1:A7",
            "stitch_mode": "",
            "output_folder": "时长提取",
            "output_name": "duration-output.xlsx",
            "remark": "1900/1904 epoch time-only parity",
        }

        try:
            python_result = legacy.run_extract_sheets(
                sources,
                [task],
                python_output,
                python_logs,
                append_existing=False,
                default_output_name="unused.xlsx",
            )
            create_rust_config(config_path, rust_output, 7)
            rust_process = subprocess.run(
                [
                    str(args.pytools_cli.resolve()),
                    "extract-sheets-area",
                    "--config",
                    str(config_path),
                    "--output-dir",
                    str(rust_output),
                    *(str(source) for source in sources),
                ],
                capture_output=True,
                text=True,
                check=False,
            )

            python_path = python_output / "时长提取" / "duration-output.xlsx"
            rust_path = rust_output / "时长提取" / "duration-output.xlsx"
            python_projection = project_workbook(python_path)
            sources_unchanged = all(
                source.read_bytes() == source_snapshots[source] for source in sources
            )
            if rust_process.returncode == 0 and rust_path.is_file():
                rust_projection = project_workbook(rust_path)
                rust_stats = json.loads(rust_process.stdout)
            else:
                rust_projection = None
                rust_stats = None
            status = (
                "passed"
                if python_result["failed_files"] == 0
                and rust_process.returncode == 0
                and rust_stats is not None
                and rust_stats["failed_files"] == 0
                and python_projection == rust_projection
                and sources_unchanged
                else "failed"
            )

            report = {
                "schema_version": 1,
                "status": status,
                "feature": "PT-7-6 area extraction time-only value classification",
                "sample_kind": (
                    "synthetic XLSX files in 1900 and 1904 epochs with midnight, rounding, "
                    "date-only, built-in time, and elapsed-duration formats"
                ),
                "source_python": "PyCode/core/feature_1_6_extract_sheets.py::run_extract_sheets",
                "rust_entry": "pytools extract-sheets-area",
                "python_stats": {
                    key: value for key, value in python_result.items() if key != "paths"
                },
                "rust_returncode": rust_process.returncode,
                "rust_stats": (
                    {key: value for key, value in rust_stats.items() if key != "paths"}
                    if rust_stats is not None
                    else None
                ),
                "python_projection": python_projection,
                "rust_error": rust_process.stderr.strip(),
                "rust_output_created": rust_path.is_file(),
                "rust_projection": rust_projection,
                "sources_unchanged": sources_unchanged,
                "scope_note": (
                    "Synthetic parity only. Real business time-only cells, row/column default "
                    "styles, unknown built-in number formats, other formats, and target-platform "
                    "delivery remain unverified. Unsupported or ambiguous metadata must fail closed."
                ),
            }
            report_path.parent.mkdir(parents=True, exist_ok=True)
            report_path.write_text(
                json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
            )
            print(json.dumps(report, ensure_ascii=False, indent=2))
            return 0 if status == "passed" else 1
        finally:
            logging.shutdown()


if __name__ == "__main__":
    raise SystemExit(main())
