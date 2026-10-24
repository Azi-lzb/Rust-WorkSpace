#!/usr/bin/env python3
"""Compare the Rust S3-F02 CLI with the read-only Python exporter on synthetic XLSX files."""

from __future__ import annotations

import argparse
import json
import math
import subprocess
import sys
import tempfile
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")


HEADERS = (
    "机构类代码", "机构类名称", "地区代码", "地区名称", "指标代码", "指标名称",
    "数据值", "上期值", "输出说明", "说明内容", "增减额（先出现）", "增减额（后出现）",
)
ROWS = (
    ("7020001", "示例村镇银行", "4401", "惠城区", "I-01", "指标一", 0.418462, 0.25, "是", "补充说明一", -0.168462, 999),
    ("12345", "普通银行", "510101", "省级", "I-02", "指标二", 10.5, 10, "否", "不应导出", 0.5, 998),
    ("12345", "普通银行", "510102", "城北区", "I-03", "指标三", 25, 20, " YES ", "补充说明三", 5, 997),
    ("99999", "地方机构", "6101", "新区", "I-04", "指标四", 40, 35, "✓", "补充说明四", 5, 996),
    ("4567", "未映射机构", "3201", "南区", "I-05", "指标五", 12, 11, "1", "补充说明五", 1, 995),
)
CONFIG_ROWS = (
    ("7020001", "4401", "精确映射机构"),
    ("12345", "510102", "精确市级机构"),
    ("12345", "*", "显式通配机构"),
    ("99999", "*", "全国通配机构"),
)


def load_baseline(source_root: Path):
    sys.path.insert(0, str(source_root / "core" / "src"))
    from base_audit.systems.s3_central_statistics.explanation_exporter import (
        export_selected_explanations,
    )

    return export_selected_explanations


def create_inputs(comparison_path: Path, config_path: Path) -> None:
    comparison = Workbook()
    cover = comparison.active
    cover.title = "封面"
    cover["A1"] = "不是比较结果"
    result = comparison.create_sheet("比较结果")
    result.append(HEADERS)
    for row in ROWS:
        result.append(row)
    comparison.save(comparison_path)

    config = Workbook()
    sheet = config.active
    sheet.title = "机构地区参照"
    sheet.append(("机构类代码", "地区代码", "机构名称"))
    for row in CONFIG_ROWS:
        sheet.append(row)
    config.save(config_path)


def color_value(color):
    if color is None:
        return None
    if color.type == "rgb":
        return ("rgb", color.rgb[-6:].upper())
    if color.type == "theme" and color.theme == 1:
        return None
    return (color.type, color.indexed if color.type == "indexed" else color.value)


def style_value(cell):
    return {
        "font": (cell.font.bold, cell.font.sz or 11, color_value(cell.font.color)),
        "fill": (cell.fill.fill_type, color_value(cell.fill.fgColor)),
        "alignment": (cell.alignment.horizontal, cell.alignment.vertical, cell.alignment.wrap_text),
        "number_format": cell.number_format,
    }


def column_width(sheet, column: str) -> float:
    index = ord(column) - ord("A") + 1
    for dimension in sheet.column_dimensions.values():
        if dimension.min is not None and dimension.max is not None:
            if dimension.min <= index <= dimension.max:
                return dimension.width
    return sheet.column_dimensions[column].width


def equal_value(left, right) -> bool:
    if isinstance(left, (int, float)) and not isinstance(left, bool):
        return isinstance(right, (int, float)) and not isinstance(right, bool) and math.isclose(
            left, right, rel_tol=0.0, abs_tol=1e-12
        )
    return left == right


def compare_workbooks(expected_path: Path, actual_path: Path) -> tuple[list[str], int]:
    expected = load_workbook(expected_path, data_only=True)
    actual = load_workbook(actual_path, data_only=True)
    differences: list[str] = []
    try:
        if expected.sheetnames != actual.sheetnames:
            differences.append("sheetnames")
        left = expected["说明文件"]
        right = actual["说明文件"]
        for row in range(1, max(left.max_row, right.max_row) + 1):
            for column in range(1, 9):
                if not equal_value(left.cell(row, column).value, right.cell(row, column).value):
                    differences.append(f"说明文件.{left.cell(row, column).coordinate}.value")
        if set(map(str, left.merged_cells.ranges)) != set(map(str, right.merged_cells.ranges)):
            differences.append("说明文件.merged_ranges")
        if left.freeze_panes != right.freeze_panes:
            differences.append("说明文件.freeze_panes")
        if left.auto_filter.ref != right.auto_filter.ref:
            differences.append("说明文件.auto_filter")
        for column in "ABCDEFGH":
            if not math.isclose(
                column_width(left, column),
                column_width(right, column),
                rel_tol=0.0,
                abs_tol=0.05,
            ):
                differences.append(f"说明文件.column_width.{column}")
        for address in ("A1", "A2", "E2", "G2", "H2", "E3", "F3", "G3"):
            if style_value(left[address]) != style_value(right[address]):
                differences.append(f"说明文件.{address}.style")

        expected_info = expected["生成信息"]
        actual_info = actual["生成信息"]
        for row in (1, 2, 4, 5):
            for column in (1, 2):
                if not equal_value(
                    expected_info.cell(row, column).value,
                    actual_info.cell(row, column).value,
                ):
                    differences.append(
                        f"生成信息.{expected_info.cell(row, column).coordinate}.value"
                    )
        for sheet in (expected_info, actual_info):
            timestamp = sheet["B3"].value
            try:
                datetime.strptime(timestamp, "%Y-%m-%d %H:%M:%S")
            except (TypeError, ValueError):
                differences.append("生成信息.B3.timestamp_format")
                break
        if expected_info["A1"].value != actual_info["A1"].value:
            differences.append("生成信息.header")
        for address in ("A1", "B1"):
            if style_value(expected_info[address]) != style_value(actual_info[address]):
                differences.append(f"生成信息.{address}.style")
        for column in "AB":
            if not math.isclose(
                column_width(expected_info, column),
                column_width(actual_info, column),
                rel_tol=0.0,
                abs_tol=0.05,
            ):
                differences.append(f"生成信息.column_width.{column}")
        return differences, left.max_row - 2
    finally:
        expected.close()
        actual.close()


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--audit-binary", required=True, type=Path)
    parser.add_argument("--source-root", required=True, type=Path)
    parser.add_argument("--report", required=True, type=Path)
    args = parser.parse_args()
    export_selected_explanations = load_baseline(args.source_root)

    result = {
        "feature": "S3-F02",
        "fixture": "synthetic",
        "baseline": "read-only Python export_selected_explanations",
        "rust_reader": "calamine",
        "rust_writer": "rust_xlsxwriter",
        "status": "failed",
    }
    try:
        with tempfile.TemporaryDirectory(prefix="audit-s3-f02-parity-") as temporary:
            root = Path(temporary)
            comparison = root / "comparison.xlsx"
            config = root / "mapping.xlsx"
            expected = root / "python-explanations.xlsx"
            actual = root / "rust-explanations.xlsx"
            create_inputs(comparison, config)
            comparison_before = comparison.read_bytes()
            config_before = config.read_bytes()
            selected = export_selected_explanations(
                comparison, expected, config_path=config
            )
            command = [
                str(args.audit_binary),
                "s3-export-explanations",
                "--comparison-file",
                str(comparison),
                "--output-workbook",
                str(actual),
                "--config-workbook",
                str(config),
            ]
            completed = subprocess.run(
                command, check=True, capture_output=True, text=True, encoding="utf-8"
            )
            report = json.loads(completed.stdout)
            differences, output_rows = compare_workbooks(expected, actual)
            if report.get("selected_rows") != selected:
                differences.append("selected_rows")
            if comparison.read_bytes() != comparison_before:
                differences.append("comparison_input_changed")
            if config.read_bytes() != config_before:
                differences.append("configuration_input_changed")
            result.update(
                {
                    "status": "passed" if not differences else "failed",
                    "selected_rows_python": selected,
                    "selected_rows_rust": report.get("selected_rows"),
                    "output_rows_compared": output_rows,
                    "differences": differences,
                    "checks": [
                        "selected rows and cell values",
                        "region code and exact/wildcard institution names",
                        "header order and comparison data preserved without rounding",
                        "sheet order, title merge, freeze pane, filter and column widths",
                        "title/header styles and numeric display formats",
                        "metadata content and timestamp format",
                        "comparison and configuration workbooks remain inputs only",
                    ],
                }
            )
    except Exception as error:  # keep the report small and avoid recording local paths
        result["error"] = f"{type(error).__name__}: {error}"

    args.report.parent.mkdir(parents=True, exist_ok=True)
    args.report.write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0 if result["status"] == "passed" else 1


if __name__ == "__main__":
    raise SystemExit(main())
