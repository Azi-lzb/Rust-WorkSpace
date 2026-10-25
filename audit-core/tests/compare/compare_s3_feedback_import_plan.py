#!/usr/bin/env python3
"""Compare the Rust S3-F03 read-only plan with the original Python importer.

This is a test helper only. It creates synthetic workbooks in a temporary
directory, calls the Rust CLI and the read-only Python baseline, and verifies
that the planned cells and conflict report match the Python write result.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

import openpyxl
from openpyxl import Workbook, load_workbook


TARGET_HEADERS = (
    "机构代码",
    "机构名称",
    "指标代码",
    "指标名称",
    "本期金额（亿元）",
    "上期值",
    "增减额（亿元）",
    "变动原因",
)
FEEDBACK_HEADERS = (
    "机构代码",
    "机构名称",
    "指标代码",
    "指标名称",
    "数据值",
    "上期值",
    "增减额（亿元）",
    "变动原因",
)


def write_target(path: Path) -> None:
    book = Workbook()
    cover = book.active
    cover.title = "封面"
    cover["A1"] = "S3-F03 synthetic parity"
    sheet = book.create_sheet("说明文件")
    sheet.merge_cells("A1:H1")
    sheet["A1"] = "xx说明"
    for column, value in enumerate(TARGET_HEADERS, 1):
        sheet.cell(2, column, value)
    rows = (
        ("6k0u", "目标机构", "A001", "指标甲", 0.418462, 1.0, -0.0002, "待机构填写"),
        ("DUP", "重复机构", "D1", "重复指标", 1.0, 1.0, 2.0, "待机构填写"),
        ("DUP", "重复机构", "D1", "重复指标", 1.0, 1.0, 2.0, "待机构填写"),
        ("MULTI", "冲突机构", "M1", "冲突指标", 3.0, 1.0, 4.0, "待机构填写"),
    )
    for row_number, row in enumerate(rows, 3):
        for column, value in enumerate(row, 1):
            sheet.cell(row_number, column, value)
    book.save(path)
    book.close()


def write_feedback(path: Path, rows: tuple[tuple[str, ...], ...]) -> None:
    book = Workbook()
    sheet = book.active
    sheet.title = "机构反馈"
    for column, value in enumerate(FEEDBACK_HEADERS, 1):
        sheet.cell(1, column, value)
    for row_number, row in enumerate(rows, 2):
        for column, value in enumerate(row, 1):
            sheet.cell(row_number, column, value)
    book.save(path)
    book.close()


def digest(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def workbook_values(path: Path) -> tuple[tuple[str, ...], ...]:
    book = load_workbook(path, data_only=False)
    try:
        result = []
        for sheet in book.worksheets:
            result.append((sheet.title,))
            for row in sheet.iter_rows():
                result.append(tuple(cell.value for cell in row))
        return tuple(result)
    finally:
        book.close()


def report_rows(path: Path) -> tuple[tuple[object, ...], ...]:
    book = load_workbook(path, data_only=True)
    try:
        sheet = book["冲突明细"]
        return tuple(tuple(row) for row in sheet.iter_rows(values_only=True))
    finally:
        book.close()


def color_hex(color: object) -> str | None:
    value = getattr(color, "rgb", None)
    return value[-6:].upper() if isinstance(value, str) else None


def report_layout(path: Path) -> dict[str, object]:
    book = load_workbook(path, data_only=True)
    try:
        sheet = book["冲突明细"]
        widths_by_column: dict[int, float] = {}
        for dimension in sheet.column_dimensions.values():
            if dimension.min is None or dimension.max is None:
                continue
            for column in range(dimension.min, dimension.max + 1):
                widths_by_column[column] = dimension.width
        if any(column not in widths_by_column for column in range(1, 11)):
            raise AssertionError("conflict report is missing an explicit width for one or more columns")
        return {
            "sheet_names": book.sheetnames,
            "freeze_panes": sheet.freeze_panes,
            "auto_filter": sheet.auto_filter.ref,
            "column_widths": [widths_by_column[column] for column in range(1, 11)],
            "header_bold": sheet["A1"].font.bold,
            "header_font_color": color_hex(sheet["A1"].font.color),
            "header_fill_color": color_hex(sheet["A1"].fill.fgColor),
        }
    finally:
        book.close()


def main() -> int:
    workspace_root = Path(__file__).resolve().parents[3]
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--audit-source-root",
        type=Path,
        default=workspace_root.parent / "3.open-data-audit",
    )
    parser.add_argument(
        "--rust-cli",
        type=Path,
        default=workspace_root / "target" / "debug" / "open-data-audit.exe",
    )
    parser.add_argument("--report-output", type=Path)
    args = parser.parse_args()
    if not args.rust_cli.is_file():
        raise SystemExit(f"Rust CLI not found: {args.rust_cli}")
    sys.path.insert(0, str(args.audit_source_root / "core" / "src"))
    from base_audit.systems.s3_central_statistics.explanation_exporter import (
        import_explanation_feedback,
    )

    with tempfile.TemporaryDirectory(prefix="audit-s3-f03-parity-") as temporary:
        root = Path(temporary)
        original_target = root / "指标说明.xlsx"
        feedback_a = root / "机构反馈A.xlsx"
        feedback_b = root / "机构反馈B.xlsx"
        write_target(original_target)
        write_feedback(
            feedback_a,
            (
                ("'6K0U", "名称有差异", "a001", "指标名称也有差异", "0.4185", "1.0", "-0.00020", "本期业务正常增加"),
                ("DUP", "重复机构", "D1", "重复指标", "1.0", "1.0", "2.0", "重复目标原因"),
                ("MULTI", "冲突机构", "M1", "冲突指标", "3.0", "1.0", "4.0", "原因甲"),
                ("MISSING", "未命中机构", "U1", "未命中指标", "9", "1", "8", "无目标原因"),
                ("IGNORED", "忽略机构", "I1", "忽略指标", "1", "1", "0", "   "),
            ),
        )
        write_feedback(
            feedback_b,
            (("MULTI", "冲突机构", "M1", "冲突指标", "3.0000", "1.0", "4.0000", "原因乙"),),
        )

        rust_target = root / "rust_target.xlsx"
        python_target = root / "python_target.xlsx"
        planned_application = root / "planned_application.xlsx"
        shutil.copy2(original_target, rust_target)
        shutil.copy2(original_target, python_target)
        candidate_path = root / "candidate.xlsx"
        plan_path = root / "plan.json"
        inputs_before = {path.name: digest(path) for path in (rust_target, feedback_a, feedback_b)}

        command = [
            str(args.rust_cli),
            "s3-plan-explanation-feedback-import",
            "--target-workbook",
            str(rust_target),
            "--feedback-workbook",
            str(feedback_a),
            "--feedback-workbook",
            str(feedback_b),
            "--candidate-output-workbook",
            str(candidate_path),
            "--plan-output",
            str(plan_path),
            "--developer-view",
        ]
        rust = subprocess.run(
            command, capture_output=True, text=True, encoding="utf-8", check=False
        )
        if rust.returncode:
            raise RuntimeError(f"Rust CLI failed ({rust.returncode}): {rust.stderr}")
        summary = json.loads(rust.stdout)
        plan = json.loads(plan_path.read_text(encoding="utf-8"))

        python_result = import_explanation_feedback(python_target, [feedback_a, feedback_b])

        # Apply the planned cells only in this temporary comparison copy.
        shutil.copy2(rust_target, planned_application)
        book = load_workbook(planned_application)
        for operation in plan["edit_plan"]["operations"]:
            book[operation["sheet"]][operation["cell"]] = operation["value"]["value"]
        book.save(planned_application)
        book.close()

        rust_conflict_path = Path(summary["conflict_report_path"])
        if not rust_conflict_path.name.startswith("rust_target_导入冲突清单_"):
            raise AssertionError("Rust conflict report name does not follow the target workbook name")
        python_conflicts = python_result.conflict_report_path
        if python_conflicts is None:
            raise AssertionError("Python baseline did not produce the expected conflict workbook")
        rust_rows = report_rows(rust_conflict_path)
        python_rows = report_rows(python_conflicts)
        rust_layout = report_layout(rust_conflict_path)
        python_layout = report_layout(python_conflicts)
        layout_matches = (
            rust_layout["sheet_names"] == python_layout["sheet_names"]
            and rust_layout["freeze_panes"] == python_layout["freeze_panes"]
            and rust_layout["auto_filter"] == python_layout["auto_filter"]
            and rust_layout["column_widths"] == python_layout["column_widths"]
            and rust_layout["header_bold"] == python_layout["header_bold"]
            and rust_layout["header_font_color"] == python_layout["header_font_color"]
            and rust_layout["header_fill_color"] == python_layout["header_fill_color"]
        )
        if workbook_values(python_target) != workbook_values(planned_application):
            raise AssertionError("Python output values differ from the Rust planned cell updates")
        if rust_rows != python_rows:
            raise AssertionError("Rust conflict rows differ from the Python conflict report")
        if not layout_matches:
            raise AssertionError(f"conflict workbook layout differs: {rust_layout} != {python_layout}")
        inputs_after = {path.name: digest(path) for path in (rust_target, feedback_a, feedback_b)}
        if inputs_before != inputs_after:
            raise AssertionError("Rust planning changed a target or feedback input")

        report = {
            "status": "passed",
            "feature": "S3-F03 read-only import planning slice",
            "baseline": "Open-Data-Audit import_explanation_feedback",
            "source_baseline": "core/src/base_audit/systems/s3_central_statistics/explanation_exporter.py",
            "platform": "Windows native",
            "python": sys.version.split()[0],
            "openpyxl": openpyxl.__version__,
            "rust_cli": args.rust_cli.name,
            "reader": summary["reader"],
            "business_engine": summary["business_engine"],
            "conflict_writer": summary["conflict_writer"],
            "counts": {
                "imported": summary["imported"],
                "unmatched": summary["unmatched"],
                "conflicts": summary["conflicts"],
                "planned_cell_updates": len(plan["edit_plan"]["operations"]),
                "python_imported": python_result.imported,
                "python_unmatched": python_result.unmatched,
                "python_conflicts": python_result.conflicts,
            },
            "planned_cell_comparison": "0 differences after applying Rust SetCellValue operations to a temporary comparison copy",
            "conflict_rows": len(rust_rows) - 1,
            "conflict_report_layout_comparison": "matched: sheet, values, freeze panes, filter, widths, header font/fill",
            "target_and_feedback_sha256_before_after": "unchanged",
            "excel_modifier_execution": "not run; target stays read-only and no S3-specific passed edit capability exists",
            "real_feedback_sample": "not available in this synthetic probe; S3-F03 remains unverified",
        }
    rendered = json.dumps(report, ensure_ascii=False, indent=2) + "\n"
    if args.report_output:
        args.report_output.parent.mkdir(parents=True, exist_ok=True)
        args.report_output.write_text(rendered, encoding="utf-8")
    print(rendered, end="")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
