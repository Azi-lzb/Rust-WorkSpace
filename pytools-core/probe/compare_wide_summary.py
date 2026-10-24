"""Compare the Rust 1-3/1-4 wide-summary menu with the read-only PyCode baseline.

The script builds a synthetic XLSX/config under a temporary directory, calls
the original Python workflow without writing to the source tree, then compares
both generated workbooks using openpyxl.
"""

from __future__ import annotations

import json
import gc
import logging
import math
import os
import shlex
import shutil
import subprocess
import sys
import tempfile
import time
import xml.etree.ElementTree as ET
from datetime import date, datetime
from importlib.metadata import version
from pathlib import Path
from zipfile import ZipFile

from openpyxl import Workbook, load_workbook
from openpyxl.comments import Comment
from openpyxl.styles import PatternFill

sys.dont_write_bytecode = True

WORKSPACE = Path(__file__).resolve().parents[2]
PYTOOLS_SOURCE = Path(r"C:\Users\AZI\Desktop\VibeCode\4.pytools-main")


def to_wsl_path(path: Path) -> str:
    drive = path.drive.rstrip(":").lower()
    if not drive:
        raise ValueError(f"expected an absolute Windows path: {path}")
    tail = path.as_posix()[2:].lstrip("/")
    return f"/mnt/{drive}/{tail}"


def from_wsl_path(path: str) -> Path:
    normalized = path.replace("\\", "/")
    if normalized.startswith("/mnt/") and len(normalized) > 6:
        drive = normalized[5].upper()
        tail = normalized[7:].replace("/", "\\")
        return Path(f"{drive}:\\{tail}")
    raise ValueError(f"Rust returned a path outside the mounted Windows workspace: {path}")


def rust_path_argument(path: Path) -> str:
    return str(path) if os.environ.get("PYTOOLS_RUST_CLI") else to_wsl_path(path)


def rust_path_result(path: str) -> Path:
    return Path(path) if os.environ.get("PYTOOLS_RUST_CLI") else from_wsl_path(path)


def run_rust_command(
    arguments: list[str], input_text: str | None = None
) -> subprocess.CompletedProcess[bytes]:
    rust_cli = os.environ.get("PYTOOLS_RUST_CLI")
    if rust_cli:
        return subprocess.run(
            [rust_cli, *arguments],
            cwd=WORKSPACE,
            check=True,
            capture_output=True,
            input=input_text.encode("utf-8") if input_text is not None else None,
        )

    workspace = to_wsl_path(WORKSPACE)
    command = (
        f"cd {shlex.quote(workspace)} && cargo run --quiet -p pytools-core --bin pytools -- "
        + " ".join(shlex.quote(argument) for argument in arguments)
    )
    return subprocess.run(
        ["wsl", "-e", "sh", "-lc", command],
        check=True,
        capture_output=True,
        input=input_text.encode("utf-8") if input_text is not None else None,
    )


def make_input(path: Path, cached_formula: bool = False) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Report"
    sheet["A1"] = "report"
    sheet.merge_cells("A2:B2")
    sheet["A2"] = "集团A"
    sheet["A3"] = "地区"
    sheet["B3"] = "存款"
    sheet.merge_cells("B3:C3")
    sheet["D3"] = "贷款"
    sheet["A4"] = "名称"
    sheet["B4"] = "余额"
    sheet["C4"] = "余额"
    sheet["D4"] = "本息"
    sheet.append([])
    sheet["A5"] = "公司甲"
    sheet["B5"] = "=8+2.25" if cached_formula else 10.25
    sheet["C5"] = 12
    sheet["D5"] = 23
    sheet["A6"] = "公司甲"
    sheet["B6"] = 14
    sheet["C6"] = 15
    sheet["D6"] = 29
    sheet["A7"] = "公司乙"
    sheet["B7"] = 4
    sheet["C7"] = 5
    sheet["D7"] = 9
    workbook.save(path)
    if cached_formula:
        patch_cached_formula(path, "B5", "8+2.25", "10.25")


def patch_cached_formula(path: Path, address: str, formula_text: str, cached_value: str) -> None:
    main_ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    ET.register_namespace("", main_ns)
    with ZipFile(path, "r") as source_archive:
        entries = [
            (info, source_archive.read(info.filename))
            for info in source_archive.infolist()
        ]

    worksheet_found = False
    patched_entries = []
    for info, payload in entries:
        if info.filename == "xl/worksheets/sheet1.xml":
            root = ET.fromstring(payload)
            cell = root.find(f".//{{{main_ns}}}c[@r='{address}']")
            if cell is None:
                raise AssertionError(f"formula cell {address} is missing from source")
            formula = cell.find(f"{{{main_ns}}}f")
            if formula is None:
                raise AssertionError(f"formula text is missing from {address}")
            formula.text = formula_text
            cached = cell.find(f"{{{main_ns}}}v")
            if cached is None:
                cached = ET.SubElement(cell, f"{{{main_ns}}}v")
            cached.text = cached_value
            payload = ET.tostring(root, encoding="utf-8", xml_declaration=True)
            worksheet_found = True
        patched_entries.append((info, payload))

    if not worksheet_found:
        raise AssertionError("source worksheet part is missing")
    with ZipFile(path, "w") as output_archive:
        for info, payload in patched_entries:
            output_archive.writestr(info, payload)

    formula_book = load_workbook(path, data_only=False, read_only=False)
    value_book = load_workbook(path, data_only=True, read_only=False)
    try:
        if formula_book["Report"][address].value != f"={formula_text}":
            raise AssertionError("formula text did not round trip in source fixture")
        if value_book["Report"][address].value != float(cached_value):
            raise AssertionError("formula cache did not round trip in source fixture")
    finally:
        formula_book.close()
        value_book.close()


def assert_cached_value_projected(paths: list[Path], expected: float) -> None:
    for path in paths:
        workbook = load_workbook(path, data_only=True, read_only=False)
        try:
            if not any(
                isinstance(cell.value, (int, float))
                and math.isclose(float(cell.value), expected, rel_tol=1e-15, abs_tol=1e-15)
                for worksheet in workbook.worksheets
                for row in worksheet.iter_rows()
                for cell in row
            ):
                raise AssertionError(f"cached formula result was not projected in {path.name}")
        finally:
            workbook.close()


def assert_blank_data_date_format(path: Path) -> None:
    workbook = load_workbook(path, data_only=False, read_only=False)
    try:
        worksheet = workbook.active
        headers = [cell.value for cell in worksheet[1]]
        date_column = headers.index("数据日期") + 1
        for row in range(2, worksheet.max_row + 1):
            cell = worksheet.cell(row, date_column)
            if cell.value is not None or cell.number_format != "yyyy/m/d":
                raise AssertionError(
                    f"blank data-date cell {cell.coordinate} must use yyyy/m/d"
                )
    finally:
        workbook.close()


def make_existing_target_seed(full_target: Path, output: Path) -> None:
    """Build a target with one duplicate row, one older row, and untouched tabs."""
    full = load_workbook(full_target, data_only=False)
    original = full["Append"]
    headers = [cell.value for cell in original[1]]
    if original.max_row < 2:
        raise AssertionError("wide-summary baseline did not create any target rows")
    first_row = [original.cell(2, column).value for column in range(1, len(headers) + 1)]
    full.close()

    workbook = Workbook()
    before = workbook.active
    before.title = "Before"
    before["A1"] = "preserve-before"
    before["A1"].fill = PatternFill("solid", fgColor="FF22AA44")
    before["A1"].comment = Comment("before-note", "probe")
    before.merge_cells("C1:D1")
    before["C1"] = "before-merge"
    before.row_dimensions[1].height = 24
    before.column_dimensions["A"].width = 19

    target = workbook.create_sheet("Append")
    for column, value in enumerate(headers, start=1):
        target.cell(1, column, value)
    for column, value in enumerate(first_row, start=1):
        target.cell(2, column, value)
    old_row = [None] * len(headers)
    old_row[headers.index("工作簿名")] = "historical.xlsx"
    old_row[headers.index("工作表名")] = "Historical"
    old_row[headers.index("数据日期")] = datetime(2023, 1, 1)
    if "集团" in headers:
        old_row[headers.index("集团")] = "旧集团"
    if "行头路径" in headers:
        old_row[headers.index("行头路径")] = "旧路径"
    target.append(old_row)
    target["A1"].fill = PatternFill("solid", fgColor="FF00FF00")
    target["A1"].comment = Comment("target-sheet-will-be-replaced", "probe")
    target.merge_cells("A50:B50")
    target["A50"] = "target-merge-will-be-replaced"
    target.row_dimensions[2].height = 40
    target.column_dimensions["B"].width = 35

    after = workbook.create_sheet("After")
    after["A1"] = "preserve-after"
    after["A1"].fill = PatternFill("solid", fgColor="FFAA4422")
    after["A1"].comment = Comment("after-note", "probe")
    after.merge_cells("C1:D1")
    after["C1"] = "after-merge"
    after.row_dimensions[1].height = 27
    after.column_dimensions["A"].width = 21
    workbook.save(output)


def make_config(path: Path, output: Path, input_dir: Path, target: Path) -> None:
    workbook = Workbook()
    global_sheet = workbook.active
    global_sheet.title = "全局配置"
    global_sheet.append(["配置项", "值"])
    for key, value in (
        ("输入目录", rust_path_argument(input_dir)),
        ("输出目录", rust_path_argument(output)),
        ("日志目录", rust_path_argument(output / "logs")),
        ("错误策略", "fail_fast"),
        ("源文件扩展名", ".xlsx"),
        ("时序规则匹配模式", "all_match"),
    ):
        global_sheet.append([key, value])

    timeline = workbook.create_sheet("时序提取规则")
    headers = [
        "是否启用", "规则名称", "工作簿关键字", "工作表关键字", "行头列", "列表头行",
        "必含列头", "必含行头", "数据起始行", "数据结束行", "数据起始列", "数据结束列",
        "跳过关键字", "目标工作簿路径", "目标工作表", "启用目标写入", "set区域", "目标去重列",
        "是否输出行序号",
    ]
    timeline.append(headers)
    timeline.append([
        "是", "WideRule", "sample", "Report", "A", "3,4", "", "", 5, "", "B", "D",
        "", rust_path_argument(target), "Append", "是", "集团@B2", "", "是",
    ])
    workbook.save(path)


def python_baseline(source: Path, output: Path, row_suffix: bool) -> tuple[Path | None, Path]:
    sys.path.insert(0, str(PYTOOLS_SOURCE / "PyCode"))
    from core.config_xlsx import GlobalConfig, TimelineRule
    from core.wide_397_398 import run_wide_summary

    output.mkdir(parents=True, exist_ok=True)
    rule = TimelineRule(
        enabled=True,
        name="WideRule",
        wb_keyword="sample",
        sheet_keyword="Report",
        row_header_col=1,
        col_header_row=3,
        row_header_cols=[1],
        row_header_col_specified=True,
        col_header_rows=[3, 4],
        data_row_start=5,
        data_col_start=2,
        data_col_end=4,
        set_items=[("集团", "B2")],
        output_row_no=True,
        target_wb_path=output / "target.xlsx",
        target_sheet="Append",
        target_write_enabled=True,
    )
    global_config = GlobalConfig(
        input_dir=source.parent,
        output_dir=output,
        log_dir=output / "logs",
        error_policy="fail_fast",
        source_exts=[".xlsx"],
    )
    try:
        summary = run_wide_summary([rule], [], global_config, row_suffix, source_paths=[source])
        return summary, output / "target.xlsx"
    finally:
        logging.shutdown()


def run_rust(config: Path, source: Path, row_suffix: bool) -> tuple[Path | None, list[Path]]:
    cli = "wide-summary-row-suffix" if row_suffix else "wide-summary"
    completed = run_rust_command(
        [cli, "--config", rust_path_argument(config), rust_path_argument(source)]
    )
    stdout = completed.stdout.decode("utf-8-sig", errors="replace").replace("\x00", "")
    stderr = completed.stderr.decode("utf-8-sig", errors="replace").replace("\x00", "")
    for line in reversed(stdout.splitlines()):
        try:
            result = json.loads(line)
            value = result.get("output")
            output = rust_path_result(value) if value else None
            plans = [rust_path_result(path) for path in result.get("target_append_plans", [])]
            return output, plans
        except json.JSONDecodeError:
            continue
    raise RuntimeError(f"Rust CLI did not return JSON: {stdout}\n{stderr}")


def run_rust_menu(
    config: Path, source: Path, row_suffix: bool
) -> tuple[Path | None, list[Path], dict[str, object]]:
    menu_choice = "3" if row_suffix else "4"
    completed = run_rust_command(
        ["menu", "--config", rust_path_argument(config), "--log-view", "developer"],
        input_text=f"1\n{menu_choice}\n{rust_path_argument(source)}\ny\n0\n",
    )
    stdout = completed.stdout.decode("utf-8-sig", errors="replace").replace("\x00", "")
    stderr = completed.stderr.decode("utf-8-sig", errors="replace").replace("\x00", "")
    for line in reversed(stdout.splitlines()):
        try:
            result = json.loads(line)
        except json.JSONDecodeError:
            continue
        if not isinstance(result, dict) or "target_append_plans" not in result:
            continue
        value = result.get("output")
        output = rust_path_result(value) if value else None
        plans = [rust_path_result(path) for path in result.get("target_append_plans", [])]
        return output, plans, result
    raise RuntimeError(
        f"Rust interactive menu did not return its wide-summary report: {stdout}\n{stderr}"
    )


def make_test_capability(plan_path: Path, capability_path: Path) -> None:
    wrapper = json.loads(plan_path.read_text(encoding="utf-8"))
    plan = wrapper["edit_plan"]
    document = {
        "engine": "umya_spreadsheet",
        "status": "passed",
        "format": "xlsx",
        "operations": sorted({operation["operation"] for operation in plan["operations"]}),
        "supported_features": plan["workbook"]["features_present"],
        "preserved": plan["preserve"],
        "platform": "synthetic-pt-1-3-1-4-umya-test",
        "evidence": {
            "sample_id": "synthetic-wide-target-sheet-replacement",
            "report_path": "pytools-core/probe/compare_wide_summary.py",
            "adapter_version": "umya-spreadsheet-3.1.1",
            "engine_version": "umya-spreadsheet-3.1.1",
            "platform": "synthetic-pt-1-3-1-4-umya-test",
        },
    }
    capability_path.write_text(json.dumps(document, ensure_ascii=False, indent=2), encoding="utf-8")


def execute_existing_target_plan(plan_path: Path, capability_path: Path) -> tuple[Path, dict[str, object]]:
    platform = "synthetic-pt-1-3-1-4-umya-test"
    completed = run_rust_command(
        [
            "dedup-append-execute",
            "--plan",
            rust_path_argument(plan_path),
            "--capabilities",
            rust_path_argument(capability_path),
            "--platform",
            platform,
            "--commit-target",
        ]
    )
    stdout = completed.stdout.decode("utf-8-sig", errors="replace").replace("\x00", "")
    stderr = completed.stderr.decode("utf-8-sig", errors="replace").replace("\x00", "")
    try:
        result = json.loads(stdout)
    except json.JSONDecodeError:
        result = None
        for line in reversed(stdout.splitlines()):
            try:
                result = json.loads(line)
                break
            except json.JSONDecodeError:
                continue
    if not isinstance(result, dict):
        raise RuntimeError(f"replacement executor did not return JSON: {stdout}\n{stderr}")
    if result.get("replacement_validation") != "passed":
        raise AssertionError(f"Rust did not report replacement validation: {result}")
    if result.get("engine") != "umya_spreadsheet":
        raise AssertionError(f"unexpected replacement engine: {result}")
    if result.get("target_commit", {}).get("operation") != "replaced":
        raise AssertionError(f"configured target was not safely replaced: {result}")
    return rust_path_result(result["output"]), result


def color_signature(color: object) -> tuple[object, ...] | None:
    if color is None:
        return None
    color_type = color.type
    if color_type == "rgb":
        return (color_type, color.rgb)
    if color_type == "indexed":
        return (color_type, color.indexed, color.tint)
    if color_type == "theme":
        return (color_type, color.theme, color.tint)
    if color_type == "auto":
        return (color_type, color.auto)
    return (color_type,)


def font_color_signature(color: object) -> tuple[object, ...] | None:
    signature = color_signature(color)
    if signature is None or signature == ("theme", 1, 0.0):
        return ("default-font-color",)
    return signature


def border_side_signature(border: object, side: str) -> tuple[object, ...]:
    value = getattr(border, side)
    if value is None:
        return (None, None)
    color = color_signature(value.color)
    if value.style is not None and color is None:
        color = ("auto", True)
    return value.style, color


def style_signature(cell: object) -> tuple[object, ...]:
    font = cell.font
    fill = cell.fill
    border = cell.border
    alignment = cell.alignment
    protection = cell.protection
    border_sides = tuple(
        border_side_signature(border, side)
        for side in ("left", "right", "top", "bottom", "diagonal", "vertical", "horizontal")
    )
    return (
        (
            font.name or "Calibri",
            float(font.sz or 11),
            font.bold,
            font.italic,
            font.vertAlign,
            font.underline,
            font.strike,
            font_color_signature(font.color),
            float(font.family or 2),
            font.charset,
            font.scheme or "minor",
            font.outline,
            font.shadow,
            font.condense,
            font.extend,
        ),
        (
            fill.fill_type,
            fill.patternType,
            color_signature(fill.fgColor),
            color_signature(fill.bgColor),
        ),
        border_sides,
        (
            border.diagonalUp,
            border.diagonalDown,
            border.outline,
        ),
        (
            alignment.horizontal,
            alignment.vertical,
            alignment.textRotation,
            alignment.wrapText,
            alignment.shrinkToFit,
            alignment.indent,
            alignment.relativeIndent,
            alignment.justifyLastLine,
            alignment.readingOrder,
        ),
        cell.number_format,
        (protection.locked, protection.hidden),
    )


def normalized(path: Path) -> dict[str, object]:
    workbook = load_workbook(path, data_only=False)
    names: list[str] = []
    rows_by_sheet: list[list[list[object]]] = []
    formats_by_sheet: list[list[str]] = []
    types_by_sheet: list[list[list[str]]] = []
    styles_by_sheet: list[list[list[tuple[object, ...]]]] = []
    dimensions_by_sheet: list[tuple[int, int]] = []
    merges_by_sheet: list[list[str]] = []
    formulas_by_sheet: list[list[tuple[str, object]]] = []
    freeze_panes_by_sheet: list[str | None] = []
    for sheet in workbook.worksheets:
        names.append(sheet.title)
        rows: list[list[object]] = []
        formats: list[str] = []
        types: list[list[str]] = []
        styles: list[list[tuple[object, ...]]] = []
        formulas: list[tuple[str, object]] = []
        for row in sheet.iter_rows():
            values = []
            row_types = []
            row_styles = []
            for cell in row:
                value = cell.value
                if isinstance(value, (datetime, date)):
                    value = value.isoformat()
                elif isinstance(value, float) and not math.isfinite(value):
                    value = str(value)
                values.append(value)
                row_types.append("blank" if cell.value is None else cell.data_type)
                row_styles.append(style_signature(cell))
                if cell.data_type == "f":
                    formulas.append((cell.coordinate, cell.value))
                if cell.row == 2:
                    formats.append(cell.number_format)
            rows.append(values)
            types.append(row_types)
            styles.append(row_styles)
        rows_by_sheet.append(rows)
        formats_by_sheet.append(formats)
        types_by_sheet.append(types)
        styles_by_sheet.append(styles)
        dimensions_by_sheet.append((sheet.max_row, sheet.max_column))
        merges_by_sheet.append(sorted(str(item) for item in sheet.merged_cells.ranges))
        formulas_by_sheet.append(formulas)
        freeze_panes_by_sheet.append(
            sheet.freeze_panes.coordinate if hasattr(sheet.freeze_panes, "coordinate") else sheet.freeze_panes
        )
    defined_names = sorted(workbook.defined_names.keys())
    workbook.close()
    return {
        "sheet_names": names,
        "rows": rows_by_sheet,
        "formats": formats_by_sheet,
        "types": types_by_sheet,
        "styles": styles_by_sheet,
        "dimensions": dimensions_by_sheet,
        "merges": merges_by_sheet,
        "formulas": formulas_by_sheet,
        "defined_names": defined_names,
        "freeze_panes": freeze_panes_by_sheet,
    }


def compare_values(expected: object, actual: object, location: str) -> None:
    if isinstance(expected, (int, float)) and isinstance(actual, (int, float)):
        if not math.isclose(float(expected), float(actual), rel_tol=1e-15, abs_tol=1e-15):
            raise AssertionError(f"{location}: {expected!r} != {actual!r}")
    elif expected != actual:
        raise AssertionError(f"{location}: {expected!r} != {actual!r}")


def assert_parity(
    expected_path: Path,
    actual_path: Path,
    compare_effective_styles: bool = True,
) -> None:
    expected = normalized(expected_path)
    actual = normalized(actual_path)
    assert expected["sheet_names"] == actual["sheet_names"], (
        expected["sheet_names"], actual["sheet_names"]
    )
    assert expected["dimensions"] == actual["dimensions"], (
        expected["dimensions"], actual["dimensions"]
    )
    assert expected["formats"] == actual["formats"], (
        expected["formats"], actual["formats"]
    )
    assert expected["merges"] == actual["merges"], (
        expected["merges"], actual["merges"]
    )
    assert expected["formulas"] == actual["formulas"], (
        expected["formulas"], actual["formulas"]
    )
    assert expected["defined_names"] == actual["defined_names"], (
        expected["defined_names"], actual["defined_names"]
    )
    assert expected["freeze_panes"] == actual["freeze_panes"], (
        expected["freeze_panes"], actual["freeze_panes"]
    )
    expected_rows = expected["rows"]
    actual_rows = actual["rows"]
    expected_types = expected["types"]
    actual_types = actual["types"]
    expected_styles = expected["styles"]
    actual_styles = actual["styles"]
    assert len(expected_rows) == len(actual_rows)
    for sheet_index, (expected_sheet, actual_sheet) in enumerate(zip(expected_rows, actual_rows)):
        assert len(expected_sheet) == len(actual_sheet)
        for row_index, (expected_row, actual_row) in enumerate(zip(expected_sheet, actual_sheet)):
            assert len(expected_row) == len(actual_row), (sheet_index, row_index, expected_row, actual_row)
            for column_index, (expected, actual) in enumerate(zip(expected_row, actual_row)):
                location = f"sheet {sheet_index} cell {row_index},{column_index}"
                compare_values(expected, actual, location)
                expected_type = expected_types[sheet_index][row_index][column_index]
                actual_type = actual_types[sheet_index][row_index][column_index]
                if expected_type != actual_type:
                    raise AssertionError(
                        f"{location} type differs: expected={expected_type!r}, actual={actual_type!r}"
                    )
                expected_style = expected_styles[sheet_index][row_index][column_index]
                actual_style = actual_styles[sheet_index][row_index][column_index]
                if compare_effective_styles and expected_style != actual_style:
                    raise AssertionError(
                        f"{location} effective style differs: "
                        f"expected={expected_style!r}, actual={actual_style!r}"
                    )


def effective_style_differences(expected_path: Path, actual_path: Path) -> list[dict[str, object]]:
    expected = normalized(expected_path)
    actual = normalized(actual_path)
    differences = []
    categories = ("font", "fill", "borders", "border_options", "alignment", "number_format", "protection")
    for sheet_index, (expected_sheet, actual_sheet) in enumerate(
        zip(expected["styles"], actual["styles"])
    ):
        for row_index, (expected_row, actual_row) in enumerate(zip(expected_sheet, actual_sheet)):
            for column_index, (expected_style, actual_style) in enumerate(
                zip(expected_row, actual_row)
            ):
                differing = [
                    category
                    for category, expected_part, actual_part in zip(
                        categories, expected_style, actual_style
                    )
                    if expected_part != actual_part
                ]
                if differing:
                    differences.append(
                        {
                            "sheet": expected["sheet_names"][sheet_index],
                            "row": row_index + 1,
                            "column": column_index + 1,
                            "attributes": differing,
                            "expected_number_format": expected_style[5],
                            "actual_number_format": actual_style[5],
                        }
                    )
    return differences


def assert_sheet_sentinels(expected_path: Path, actual_path: Path, names: tuple[str, ...]) -> None:
    expected = load_workbook(expected_path, data_only=False)
    actual = load_workbook(actual_path, data_only=False)
    for name in names:
        left = expected[name]
        right = actual[name]
        assert left["A1"].value == right["A1"].value
        assert left["A1"].fill.fgColor.rgb == right["A1"].fill.fgColor.rgb
        assert (left["A1"].comment.text if left["A1"].comment else None) == (
            right["A1"].comment.text if right["A1"].comment else None
        )
        assert set(map(str, left.merged_cells.ranges)) == set(map(str, right.merged_cells.ranges))
        assert left.row_dimensions[1].height == right.row_dimensions[1].height
        assert left.column_dimensions["A"].width == right.column_dimensions["A"].width
    expected.close()
    actual.close()


def assert_replacement_drops(seed_path: Path, output_path: Path) -> None:
    seed = load_workbook(seed_path, data_only=False)
    output = load_workbook(output_path, data_only=False)
    source = seed["Append"]
    replaced = output["Append"]
    assert source["A1"].fill.fgColor.rgb == "FF00FF00"
    assert source["A1"].comment is not None
    assert "A50:B50" in {str(item) for item in source.merged_cells.ranges}
    assert source.row_dimensions[2].height == 40
    assert source.column_dimensions["B"].width == 35
    assert replaced["A1"].fill.fgColor.rgb != "FF00FF00"
    assert replaced["A1"].comment is None
    assert "A50:B50" not in {str(item) for item in replaced.merged_cells.ranges}
    assert replaced.row_dimensions[2].height != 40
    assert replaced.column_dimensions["B"].width != 35
    seed.close()
    output.close()


def run_case(
    root: Path,
    row_suffix: bool,
    *,
    source_filename: str = "sample_2024-06-30.xlsx",
    cached_formula: bool = False,
) -> dict[str, object]:
    case_name = (
        "no_date_formula_cache"
        if cached_formula
        else "suffix"
        if row_suffix
        else "no_suffix"
    )
    case = root / case_name
    case.mkdir()
    source = case / source_filename
    make_input(source, cached_formula=cached_formula)
    python_output = case / "python_output"
    rust_output = case / "rust_output"
    config = case / "config.xlsx"
    make_config(config, rust_output, source.parent, rust_output / "target.xlsx")
    source_before = source.read_bytes()
    config_before = config.read_bytes()
    baseline, python_target = python_baseline(source, python_output, row_suffix)
    actual, plans, menu_report = run_rust_menu(config, source, row_suffix)
    if baseline is None or actual is None:
        raise AssertionError(f"wide summary did not produce output: baseline={baseline}, rust={actual}")
    assert_parity(baseline, actual)
    if len(plans) != 1:
        raise AssertionError(f"expected one target append plan, got {plans}")
    wrapper = json.loads(plans[0].read_text(encoding="utf-8"))
    candidate = rust_path_result(wrapper["candidate"])
    committed_target = rust_path_result(wrapper["target"])
    executions = menu_report.get("target_append_executions", [])
    if (
        menu_report.get("target_append_status") != "succeeded"
        or len(executions) != 1
        or executions[0].get("status") != "succeeded"
    ):
        raise AssertionError(f"interactive target append did not succeed: {executions}")
    commit_execution = executions[0]
    if commit_execution.get("operation") != "created":
        raise AssertionError(
            f"interactive menu did not create the configured target: {commit_execution}"
        )
    if not candidate.is_file() or not committed_target.is_file():
        raise AssertionError("interactive menu did not retain candidate and committed target")
    run_log_path = menu_report.get("run_log_path")
    if not run_log_path:
        raise AssertionError("developer-view menu report did not include its run history")
    history = json.loads(rust_path_result(run_log_path).read_text(encoding="utf-8"))
    target_events = [
        event
        for event in history.get("events", [])
        if event.get("phase") == "target_append_execute"
    ]
    if len(target_events) != 1:
        raise AssertionError(f"expected one target execution history event: {target_events}")
    target_event = target_events[0]
    if (
        target_event.get("status") != "succeeded"
        or target_event.get("actual_engine") != "rust_xlsxwriter"
        or not isinstance(target_event.get("stage_elapsed_ms"), int)
    ):
        raise AssertionError(f"unexpected target execution history event: {target_event}")
    assert_parity(python_target, candidate)
    assert_parity(python_target, committed_target)
    if cached_formula:
        assert_cached_value_projected(
            [baseline, actual, python_target, candidate, committed_target], 10.25
        )
        for output_path in [baseline, actual, python_target, candidate, committed_target]:
            assert_blank_data_date_format(output_path)
    if source.read_bytes() != source_before or config.read_bytes() != config_before:
        raise AssertionError("interactive Rust/Python run modified source workbook or config")
    if committed_target.read_bytes() != candidate.read_bytes():
        raise AssertionError("new target commit differs from candidate")
    projection = normalized(actual)
    result = {
        "row_suffix": row_suffix,
        "cached_formula_value_projected": cached_formula,
        "blank_data_date_format": "yyyy/m/d" if cached_formula else "not_applicable",
        "summary_status": "passed",
        "target_append_status": "passed",
        "summary_sheet_names": projection["sheet_names"],
        "summary_dimensions": projection["dimensions"],
        "summary_cell_count": sum(
            len(row) for sheet in projection["rows"] for row in sheet
        ),
        "summary_effective_styles": "exact",
        "entry_route": "pytools menu -> 1 timeline -> 3 row-suffix or 4 no-suffix -> confirm",
        "target_commit_operation": commit_execution["operation"],
        "target_candidate_matches_committed_bytes": True,
        "target_execution_history": {
            "phase": target_event["phase"],
            "status": target_event["status"],
            "actual_engine": target_event["actual_engine"],
            "stage_elapsed_ms_recorded": True,
        },
        "target_execution_summary": commit_execution["execution_summary"],
    }
    print(json.dumps(result, ensure_ascii=False))
    return result


def run_existing_target_case(root: Path, row_suffix: bool) -> dict[str, object]:
    case = root / ("existing_suffix" if row_suffix else "existing_no_suffix")
    case.mkdir()
    source = case / "sample_2024-06-30.xlsx"
    make_input(source)
    seed_output = case / "seed_output"
    _, full_target = python_baseline(source, seed_output, row_suffix)
    seed_target = case / "target_seed.xlsx"
    make_existing_target_seed(full_target, seed_target)
    seed_bytes = seed_target.read_bytes()

    python_output = case / "python_output"
    rust_output = case / "rust_output"
    python_output.mkdir()
    rust_output.mkdir()
    python_target = python_output / "target.xlsx"
    rust_target = rust_output / "target.xlsx"
    python_target.write_bytes(seed_bytes)
    rust_target.write_bytes(seed_bytes)
    config = case / "config.xlsx"
    make_config(config, rust_output, source.parent, rust_target)
    source_before = source.read_bytes()
    config_before = config.read_bytes()

    python_summary, expected = python_baseline(source, python_output, row_suffix)
    actual_summary, plans = run_rust(config, source, row_suffix)
    if python_summary is None or actual_summary is None:
        raise AssertionError("existing-target wide-summary did not produce both summary workbooks")
    assert_parity(python_summary, actual_summary)
    if len(plans) != 1:
        raise AssertionError(f"expected one existing-target plan, got {plans}")
    plan_path = plans[0]
    wrapper = json.loads(plan_path.read_text(encoding="utf-8"))
    if wrapper.get("replacement_workbook") is None:
        raise AssertionError("PT-1-3/PT-1-4 plan did not contain the target-sheet replacement payload")
    if wrapper["edit_plan"].get("drop_allowed", {}).get("Append") is None:
        raise AssertionError("replacement plan did not scope intentional drops to the replaced worksheet")
    capability_path = case / "temporary-test-capability.json"
    make_test_capability(plan_path, capability_path)
    candidate, execution = execute_existing_target_plan(plan_path, capability_path)
    assert_parity(expected, candidate, compare_effective_styles=False)
    style_differences = effective_style_differences(expected, candidate)
    assert_sheet_sentinels(expected, candidate, ("Before", "After"))
    assert_replacement_drops(seed_target, expected)
    assert_replacement_drops(seed_target, candidate)
    if rust_target.read_bytes() != candidate.read_bytes():
        raise AssertionError("configured target bytes differ from the validated candidate")
    if rust_target.read_bytes() == seed_bytes:
        raise AssertionError("Rust execution left the configured target unchanged")
    if source.read_bytes() != source_before or config.read_bytes() != config_before:
        raise AssertionError("Rust/Python parity run modified the source workbook or configuration")
    candidate_projection = normalized(candidate)
    candidate_names = candidate_projection["sheet_names"]
    candidate_rows = candidate_projection["rows"]
    result = {
        "row_suffix": row_suffix,
        "summary_status": "passed",
        "existing_target_replacement_status": "passed",
        "engine": execution["engine"],
        "target_commit_operation": execution["target_commit"]["operation"],
        "execution_summary": execution["execution_summary"],
        "candidate_artifacts_retained": False,
        "candidate_matches_committed_target_bytes": rust_target.read_bytes() == candidate.read_bytes(),
        "compared_sheet_order": candidate_names,
        "compared_cell_count": sum(len(row) for sheet in candidate_rows for row in sheet),
        "target_sheet_style_status": "exact" if not style_differences else "mismatch_detected",
        "target_sheet_style_differences": style_differences,
        "source_input_and_config_unchanged": True,
        "configured_target_bytes_unchanged": False,
        "configured_target_committed": True,
        "configured_target_changed_from_seed": rust_target.read_bytes() != seed_bytes,
        "intentional_target_sheet_drops": [
            "styles", "comments", "merged_ranges", "row_dimensions", "column_dimensions"
        ],
        "untouched_sheets": ["Before", "After"],
        "untouched_sheet_checks": ["values", "style_fill", "comments", "merges", "row_height", "column_width"],
        "temporary_test_capability_only": True,
    }
    print(json.dumps(result, ensure_ascii=False))
    return result


def cleanup_probe_directory(path: Path) -> None:
    """Retry transient Windows/WSL file-handle release before temp cleanup."""
    gc.collect()
    for attempt in range(8):
        try:
            shutil.rmtree(path)
            return
        except PermissionError:
            if attempt == 7:
                raise
            gc.collect()
            time.sleep(0.25 * (attempt + 1))


def main() -> int:
    if not PYTOOLS_SOURCE.is_dir():
        raise FileNotFoundError(f"PyCode source not found: {PYTOOLS_SOURCE}")
    with tempfile.TemporaryDirectory(prefix="pytools-wide-probe-", ignore_cleanup_errors=True) as temp:
        root = Path(temp)
        summary_cases = [
            run_case(root, False),
            run_case(root, True),
            run_case(
                root,
                False,
                source_filename="sample_without_date.xlsx",
                cached_formula=True,
            ),
        ]
        existing_cases = [
            run_existing_target_case(root, False),
            run_existing_target_case(root, True),
        ]
        report = {
            "schema_version": 1,
            "suite": "pytools.pt_1_3_1_4_existing_target_replacement_parity",
            "status": "passed_with_existing_target_style_mismatch"
            if any(case["target_sheet_style_status"] != "exact" for case in existing_cases)
            else "passed_synthetic_parity",
            "environment": {
                "python": sys.version.split()[0],
                "python_packages": {
                    "numpy": version("numpy"),
                    "pandas": version("pandas"),
                    "openpyxl": version("openpyxl"),
                },
                "python_baseline": str(PYTOOLS_SOURCE),
                "rust_execution": (
                    "Windows native pytools menu and CLI"
                    if os.environ.get("PYTOOLS_RUST_CLI")
                    else "WSL2 cargo run via pytools menu and CLI; Windows Python executes baseline and comparisons"
                ),
                "sample_kind": "synthetic XLSX only",
            },
            "comparison_contract": {
                "worksheet_names_order_values_and_types": "exact except numeric values compared with relative and absolute tolerance 1e-15",
                "effective_cell_styles": "exact across every cell in generated summary workbooks, including font, fill, border, alignment, number format, and protection",
                "formula_text_merges_dimensions_defined_names_freeze_panes": "exact; cached source formula values are projected as values, matching PyCode; generated summary output contains no formulas, merges, defined names, or frozen panes",
                "data_date_number_formats": "exact across the generated summary worksheet",
                "new_target_entry": "pytools menu -> timeline 1 -> wide summary 3/4 -> explicit confirmation",
                "new_target_execution_history": "developer-view history records successful target_append_execute, rust_xlsxwriter and stage duration",
                "existing_target_behavior": "rebuild target worksheet from existing rows plus accepted non-duplicate rows; preserve untouched sheets",
                "target_sheet_replacement": "values, types, dimensions, and required untouched-sheet properties are checked; effective cell-style differences are reported and block acceptance",
                "untouched_sheet_properties": "selected values, A1 fill/comment, merges, row height, and column width",
                "original_target_input": "configured target is replaced only after capability execution, candidate validation, SHA-256 check, atomic sibling publication, and Calamine reopen; source and config remain unchanged",
                "test_capability": "temporary per-run capability; not registered in workspace production capabilities",
            },
            "summary_workbook_cases": summary_cases,
            "existing_target_cases": existing_cases,
            "commands": {
                "probe": "<PyTools Python 3.12.10 venv> pytools-core/probe/compare_wide_summary.py",
                "rust_build": "cargo build --offline --locked -p pytools-core --bin pytools",
            },
            "limitations": [
                "No sanitized real business sample was available.",
                "Temporary capability evidence is test-only and does not register a production route.",
                "麒麟/UOS glibc 2.28 deb remains unverified.",
                "The cached-formula scenario validates reading and projecting a stored formula result; it does not perform or validate formula recalculation.",
            ],
        }
        report_path = WORKSPACE / "pytools-core" / "probe" / "reports" / "wide_summary_existing_target_parity_20261023.json"
        report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        summary_report = {
            "schema_version": 1,
            "suite": "pytools.timeline.wide_summary_python_rust_parity",
            "date": date.today().isoformat(),
            "status": "passed_synthetic_parity",
            "source_python_files": [
                "PyCode/core/wide_397_398.py",
                "PyCode/core/io_excel.py",
            ],
            "input": {
                "format": "xlsx",
                "kind": "synthetic",
                "features": [
                    "merged set-region label",
                    "multi-row and merged column headers",
                    "duplicate dynamic headers",
                    "duplicate row labels",
                    "date in the input filename",
                    "configured target worksheet",
                    "cached formula value with no parsable input date",
                ],
            },
            "environment": report["environment"],
            "comparison": {
                "worksheet_names_order_values_and_types": report["comparison_contract"]["worksheet_names_order_values_and_types"],
                "effective_cell_styles": report["comparison_contract"]["effective_cell_styles"],
                "formula_text_merges_dimensions_defined_names_freeze_panes": report["comparison_contract"]["formula_text_merges_dimensions_defined_names_freeze_panes"],
                "data_date_number_formats": report["comparison_contract"]["data_date_number_formats"],
                "new_target_append": "exact after the real Rust menu generated, validated, confirmed and committed the candidate",
            },
            "scenarios": summary_cases,
            "limitations": report["limitations"],
        }
        summary_report_path = WORKSPACE / "pytools-core" / "probe" / "reports" / "wide_summary_parity_20261023.json"
        summary_report_path.write_text(
            json.dumps(summary_report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
        )
        print(
            json.dumps(
                {
                    "report": str(summary_report_path),
                    "existing_target_report": str(report_path),
                    "summary_cases": len(summary_cases),
                    "existing_target_cases": len(existing_cases),
                },
                ensure_ascii=False,
            )
        )
        cleanup_probe_directory(root)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
