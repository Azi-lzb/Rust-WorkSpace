#!/usr/bin/env python3
"""Compare an Unverified Rust ImportRange candidate with Audit Python copy semantics.

The helper creates only temporary outputs. It never writes the supplied template
or audit workbook and emits aggregate counts only; no cell values or formulas are
persisted in the report.
"""

from __future__ import annotations

import argparse
import json
import math
import os
import shlex
import shutil
import subprocess
import sys
import tempfile
import time
import zipfile
from pathlib import Path
from types import SimpleNamespace
from xml.etree import ElementTree

from openpyxl import Workbook, load_workbook
from openpyxl.comments import Comment
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.styles import Alignment, Font, PatternFill, Protection
from openpyxl.utils.cell import range_boundaries


ROOT = Path(__file__).resolve().parents[3]
MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"


class CandidateReadbackError(Exception):
    pass


def wsl_path(path: Path) -> str:
    drive = path.drive.rstrip(":").lower()
    if len(drive) != 1:
        raise ValueError("candidate path must be on a mounted Windows drive")
    relative = path.as_posix().split("/", 1)[1]
    return f"/mnt/{drive}/{relative}"


def make_synthetic_pair(directory: Path) -> tuple[Path, Path, list[dict[str, str]]]:
    template_path = directory / "template.xlsx"
    audit_path = directory / "audit.xlsx"
    ranges = ["A1:B2", "D1:E2", "G1:H2", "J1:K2"]

    template = Workbook()
    source = template.active
    source.title = "审核表"
    for index, address in enumerate(ranges, start=1):
        min_col, min_row, max_col, max_row = range_boundaries(address)
        for row in range(min_row, max_row + 1):
            for column in range(min_col, max_col + 1):
                cell = source.cell(row, column)
                cell.value = f"source-{index}-{row}-{column}"
                cell.font = Font(bold=(row == min_row), color="FF223344")
                cell.fill = PatternFill("solid", fgColor="FFEFEFEF")
                cell.alignment = Alignment(horizontal="center", wrap_text=True)
                cell.protection = Protection(locked=(column % 2 == 0), hidden=False)
                cell.number_format = "0.00"
        top_left = source.cell(min_row, min_col)
        top_left.value = f"=ROW()+COLUMN()+{index}"
        top_left.comment = Comment(f"synthetic-rule-{index}", "baseline")
    source["B2"].comment = Comment("synthetic-value-comment", "baseline")
    template.defined_names.add(
        DefinedName(
            "校验区域",
            attr_text=",".join(f"'审核表'!{address}" for address in ranges),
        )
    )
    template.save(template_path)
    template.close()

    audit = load_workbook(template_path, data_only=False, keep_links=True)
    destination = audit["审核表"]
    destination["M1"] = "target-only-sentinel"
    destination["M1"].comment = Comment("untouched-comment", "target")
    for index, address in enumerate(ranges, start=1):
        min_col, min_row, max_col, max_row = range_boundaries(address)
        for row in range(min_row, max_row + 1):
            for column in range(min_col, max_col + 1):
                destination.cell(row, column).value = f"old-{index}-{row}-{column}"
    untouched = audit.create_sheet("Untouched")
    untouched["A1"] = "preserve-sheet"
    untouched["A1"].font = Font(italic=True)
    audit.save(audit_path)
    audit.close()

    return template_path, audit_path, [
        {"sheet_name": "审核表", "address": address} for address in ranges
    ]


def resolve_real_ranges(source_root: Path, template_path: Path, mapping_path: Path | None):
    sys.path.insert(0, str(source_root / "core" / "src"))
    from base_audit.name_config import (
        FORMULA_COPY_FUNCTION,
        ISSUE_EXTRACT_FUNCTION,
        STRUCTURE_COMPARE_FUNCTION,
        features_of_type,
        load_feature_mappings,
    )
    from base_audit.native.openpyxl_workbook import read_template

    mappings = load_feature_mappings(mapping_path, template_path)
    definition = read_template(
        template_path,
        formula_mappings=features_of_type(mappings, FORMULA_COPY_FUNCTION),
        structure_mappings=features_of_type(mappings, STRUCTURE_COMPARE_FUNCTION),
        extraction_mappings=features_of_type(mappings, ISSUE_EXTRACT_FUNCTION),
    )
    if definition.structured or len(definition.copy_ranges) != 4:
        raise ValueError("real candidate scope must be four named formula ranges")
    return definition


def color_signature(color):
    if color is None:
        return None
    return (color.type, color.rgb, color.indexed, color.theme, color.tint, color.auto)


def style_signature(cell):
    font = cell.font
    fill = cell.fill
    border = cell.border
    alignment = cell.alignment
    protection = cell.protection
    return (
        cell.number_format,
        (font.name, font.sz, font.b, font.i, font.u, font.strike, color_signature(font.color)),
        (fill.fill_type, color_signature(fill.fgColor), color_signature(fill.bgColor)),
        tuple(
            (name, getattr(getattr(border, name), "style", None), color_signature(getattr(getattr(border, name), "color", None)))
            for name in ("left", "right", "top", "bottom", "diagonal", "vertical", "horizontal")
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
        (protection.locked, protection.hidden),
    )


def cell_signature(cell):
    comment = None if cell.comment is None else (cell.comment.text, cell.comment.author)
    hyperlink = None
    if cell.hyperlink is not None:
        hyperlink = (
            cell.hyperlink.target,
            cell.hyperlink.location,
            cell.hyperlink.display,
            cell.hyperlink.tooltip,
        )
    return cell.value, cell.data_type, style_signature(cell), comment, hyperlink


def cell_mismatch_counts(expected_path: Path, actual_path: Path, ranges):
    expected = load_workbook(expected_path, data_only=False, keep_links=True)
    actual = load_workbook(actual_path, data_only=False, keep_links=True)
    counts = {"range_cells": 0, "range_mismatches": 0, "outside_cells": 0, "outside_mismatches": 0}
    range_cells: dict[str, set[tuple[int, int]]] = {}
    for item in ranges:
        min_col, min_row, max_col, max_row = range_boundaries(item["address"].replace("$", ""))
        cells = {(row, col) for row in range(min_row, max_row + 1) for col in range(min_col, max_col + 1)}
        range_cells.setdefault(item["sheet_name"], set()).update(cells)
        left = expected[item["sheet_name"]]
        right = actual[item["sheet_name"]]
        for row, col in cells:
            counts["range_cells"] += 1
            try:
                left_signature = cell_signature(left.cell(row, col))
            except Exception as error:
                raise CandidateReadbackError(f"python_range_cell_{type(error).__name__}") from None
            try:
                right_signature = cell_signature(right.cell(row, col))
            except Exception as error:
                raise CandidateReadbackError(f"rust_range_cell_{type(error).__name__}") from None
            counts["range_mismatches"] += int(left_signature != right_signature)

    for sheet_name in sorted(set(expected.sheetnames) | set(actual.sheetnames)):
        if sheet_name not in expected.sheetnames or sheet_name not in actual.sheetnames:
            counts["outside_mismatches"] += 1
            continue
        left = expected[sheet_name]
        right = actual[sheet_name]
        coordinates = set(left._cells) | set(right._cells)
        for row, col in coordinates:
            if (row, col) in range_cells.get(sheet_name, set()):
                continue
            counts["outside_cells"] += 1
            try:
                left_signature = cell_signature(left.cell(row, col))
            except Exception as error:
                raise CandidateReadbackError(f"python_outside_cell_{type(error).__name__}") from None
            try:
                right_signature = cell_signature(right.cell(row, col))
            except Exception as error:
                raise CandidateReadbackError(f"rust_outside_cell_{type(error).__name__}") from None
            counts["outside_mismatches"] += int(left_signature != right_signature)

    expected.close()
    actual.close()
    return counts


def workbook_metadata(path: Path):
    workbook = load_workbook(path, data_only=False, keep_links=True)
    result = {
        "sheet_names": tuple(workbook.sheetnames),
        "defined_names": tuple(sorted((item.name, item.attr_text, item.localSheetId, item.hidden) for item in workbook.defined_names.values())),
        "sheets": {},
    }
    for sheet in workbook.worksheets:
        validations = tuple(
            sorted(
                (
                    str(item.sqref), item.type, item.operator, item.formula1, item.formula2,
                    item.allow_blank, item.showErrorMessage, item.showInputMessage,
                )
                for item in sheet.data_validations.dataValidation
            )
        )
        conditional = tuple(
            sorted((str(group.sqref), repr(group.rules)) for group in sheet.conditional_formatting)
        )
        dimensions = (
            tuple(sorted((key, value.height, value.hidden, value.outlineLevel, value.collapsed) for key, value in sheet.row_dimensions.items())),
            tuple(sorted((key, value.width, value.hidden, value.outlineLevel, value.collapsed) for key, value in sheet.column_dimensions.items())),
        )
        result["sheets"][sheet.title] = (
            tuple(sorted(str(item) for item in sheet.merged_cells.ranges)),
            validations,
            conditional,
            dimensions,
            sheet.freeze_panes,
            sheet.print_area,
            sheet.print_title_rows,
            sheet.print_title_cols,
            sheet.page_setup.orientation,
            sheet.page_setup.paperSize,
            sheet.page_setup.fitToWidth,
            sheet.page_setup.fitToHeight,
            sheet.sheet_properties.pageSetUpPr.fitToPage if sheet.sheet_properties.pageSetUpPr else None,
            (sheet.page_margins.left, sheet.page_margins.right, sheet.page_margins.top, sheet.page_margins.bottom),
            (str(sheet.oddHeader), str(sheet.oddFooter), str(sheet.evenHeader), str(sheet.evenFooter)),
            sheet.sheet_state,
        )
    workbook.close()
    return result


def calculation_flags(path: Path):
    with zipfile.ZipFile(path) as archive:
        root = ElementTree.fromstring(archive.read("xl/workbook.xml"))
    node = root.find(f"{{{MAIN_NS}}}calcPr")
    if node is None:
        return {"full_calc_on_load": None, "force_full_calc": None}
    return {
        "full_calc_on_load": node.get("fullCalcOnLoad"),
        "force_full_calc": node.get("forceFullCalc"),
    }


def style_reference_summary(path: Path):
    with zipfile.ZipFile(path) as archive:
        styles = ElementTree.fromstring(archive.read("xl/styles.xml"))
        cell_xfs = styles.find(f"{{{MAIN_NS}}}cellXfs")
        style_count = len(cell_xfs) if cell_xfs is not None else 0
        fonts = styles.find(f"{{{MAIN_NS}}}fonts")
        fills = styles.find(f"{{{MAIN_NS}}}fills")
        borders = styles.find(f"{{{MAIN_NS}}}borders")
        font_count = len(fonts) if fonts is not None else 0
        fill_count = len(fills) if fills is not None else 0
        border_count = len(borders) if borders is not None else 0
        custom_formats = {
            int(item.get("numFmtId"))
            for item in styles.findall(f".//{{{MAIN_NS}}}numFmt")
            if item.get("numFmtId")
        }
        references = 0
        invalid = 0
        max_index = -1
        invalid_sheets = set()
        invalid_components = {"font": 0, "fill": 0, "border": 0, "custom_number_format": 0}
        for name in archive.namelist():
            if not (name.startswith("xl/worksheets/") and name.endswith(".xml") and "/_rels/" not in name):
                continue
            root = ElementTree.fromstring(archive.read(name))
            for cell in root.findall(f".//{{{MAIN_NS}}}c"):
                raw_index = cell.get("s")
                if raw_index is None:
                    continue
                index = int(raw_index)
                references += 1
                max_index = max(max_index, index)
                if index >= style_count:
                    invalid += 1
                    invalid_sheets.add(name)
                    continue
                style = cell_xfs[index]
                for component, attribute, maximum in (
                    ("font", "fontId", font_count),
                    ("fill", "fillId", fill_count),
                    ("border", "borderId", border_count),
                ):
                    component_index = int(style.get(attribute, "0"))
                    if component_index >= maximum:
                        invalid_components[component] += 1
                number_format = int(style.get("numFmtId", "0"))
                if number_format >= 164 and number_format not in custom_formats:
                    invalid_components["custom_number_format"] += 1
    return {
        "cell_xf_count": style_count,
        "cell_style_reference_count": references,
        "invalid_cell_style_reference_count": invalid,
        "max_cell_style_index": max_index,
        "sheets_with_invalid_style_references": len(invalid_sheets),
        "invalid_style_component_reference_counts": invalid_components,
    }


def formula_cache_count(path: Path) -> int:
    count = 0
    with zipfile.ZipFile(path) as archive:
        for name in archive.namelist():
            if name.startswith("xl/worksheets/") and name.endswith(".xml") and "/_rels/" not in name:
                root = ElementTree.fromstring(archive.read(name))
                for cell in root.findall(f".//{{{MAIN_NS}}}c"):
                    formula = cell.find(f"{{{MAIN_NS}}}f")
                    value = cell.find(f"{{{MAIN_NS}}}v")
                    if formula is not None and value is not None and value.text not in (None, ""):
                        count += 1
    return count


def formula_cache_snapshot(path: Path):
    formula_book = load_workbook(path, data_only=False, read_only=True, keep_links=False)
    cached_book = load_workbook(path, data_only=True, read_only=True, keep_links=False)
    values = {}
    try:
        for formula_sheet in formula_book.worksheets:
            cached_sheet = cached_book[formula_sheet.title]
            for row in formula_sheet.iter_rows():
                for cell in row:
                    if cell.data_type == "f":
                        values[(formula_sheet.title, cell.coordinate)] = (
                            cell.value,
                            cached_sheet[cell.coordinate].value,
                        )
    finally:
        formula_book.close()
        cached_book.close()
    return values


def cached_value_equal(left, right) -> bool:
    if isinstance(left, bool) or isinstance(right, bool):
        return type(left) is type(right) and left == right
    if isinstance(left, (int, float)) and isinstance(right, (int, float)):
        return math.isclose(float(left), float(right), rel_tol=1e-15, abs_tol=0.0)
    return type(left) is type(right) and left == right


def recalculate_and_compare(
    sidecar_path: Path,
    python_output: Path,
    rust_output: Path,
    directory: Path,
    expected_cache_values: dict[tuple[str, str], object] | None = None,
):
    started = time.perf_counter()
    python_recalculated = directory / "python_recalculated.xlsx"
    rust_recalculated = directory / "rust_recalculated.xlsx"
    shutil.copy2(python_output, python_recalculated)

    from base_audit.engines.libreoffice import LibreOfficeCalculator
    calculator = LibreOfficeCalculator()

    try:
        python_result = calculator.recalculate_many(
            (python_recalculated,),
            batch_size=8,
        )
    except Exception as error:
        raise CandidateReadbackError(f"python_recalculation_{type(error).__name__}") from None

    try:
        sidecar_result = subprocess.run(
            [
                str(sidecar_path),
                "recalculate",
                "--input",
                str(rust_output),
                "--output",
                str(rust_recalculated),
                "--engine",
                "lo",
            ],
            timeout=calculator.timeout_seconds,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            check=False,
        )
    except subprocess.TimeoutExpired:
        raise CandidateReadbackError("sidecar_timeout") from None
    except OSError as error:
        raise CandidateReadbackError(f"sidecar_start_{type(error).__name__}") from None
    if sidecar_result.returncode != 0:
        raise CandidateReadbackError(f"sidecar_exit_{sidecar_result.returncode}")

    try:
        sidecar_report = json.loads(sidecar_result.stdout)
    except (json.JSONDecodeError, TypeError):
        raise CandidateReadbackError("sidecar_invalid_json") from None
    if not isinstance(sidecar_report, dict):
        raise CandidateReadbackError("sidecar_invalid_response")
    validation = sidecar_report.get("validation", {})
    if not isinstance(validation, dict):
        raise CandidateReadbackError("sidecar_invalid_validation")
    formula_count = validation.get("formula_count")
    formula_cache_cells = validation.get("formula_cache_cells")
    if (
        sidecar_report.get("status") != "recalculated"
        or not isinstance(sidecar_report.get("engine"), str)
        or "LibreOffice" not in sidecar_report["engine"]
        or validation.get("formula_text_preserved") is not True
        or validation.get("cache_nodes_complete") is not True
        or not isinstance(formula_count, int)
        or isinstance(formula_count, bool)
        or formula_count <= 0
        or formula_cache_cells != formula_count
        or not rust_recalculated.is_file()
        or rust_recalculated.stat().st_size == 0
    ):
        raise CandidateReadbackError("sidecar_validation_failed")

    python_values = formula_cache_snapshot(python_recalculated)
    rust_values = formula_cache_snapshot(rust_recalculated)
    python_keys = set(python_values)
    rust_keys = set(rust_values)
    formula_set_difference_count = len(python_keys ^ rust_keys)
    formula_text_mismatch_count = 0
    cache_value_mismatch_count = 0
    for key in python_keys & rust_keys:
        python_formula, python_cache = python_values[key]
        rust_formula, rust_cache = rust_values[key]
        formula_text_mismatch_count += int(python_formula != rust_formula)
        cache_value_mismatch_count += int(not cached_value_equal(python_cache, rust_cache))

    python_expected_cache_mismatch_count = 0
    sidecar_expected_cache_mismatch_count = 0
    if expected_cache_values is not None:
        python_expected_cache_mismatch_count = sum(
            key not in python_values
            or not cached_value_equal(python_values[key][1], expected_value)
            for key, expected_value in expected_cache_values.items()
        )
        sidecar_expected_cache_mismatch_count = sum(
            key not in rust_values
            or not cached_value_equal(rust_values[key][1], expected_value)
            for key, expected_value in expected_cache_values.items()
        )
    expected_cache_mismatch_count = (
        python_expected_cache_mismatch_count + sidecar_expected_cache_mismatch_count
    )

    mismatch_count = (
        formula_set_difference_count
        + formula_text_mismatch_count
        + cache_value_mismatch_count
        + expected_cache_mismatch_count
    )
    return {
        "status": "matched" if mismatch_count == 0 else "mismatch",
        "python_engine": python_result.engine_display,
        "sidecar_engine": sidecar_report["engine"],
        "formula_count_python": len(python_keys),
        "formula_count_sidecar": len(rust_keys),
        "python_cached_values_nonempty": sum(value is not None for _, value in python_values.values()),
        "sidecar_cached_values_nonempty": sum(value is not None for _, value in rust_values.values()),
        "formula_set_difference_count": formula_set_difference_count,
        "formula_text_mismatch_count": formula_text_mismatch_count,
        "cache_value_mismatch_count": cache_value_mismatch_count,
        "expected_cache_cells": len(expected_cache_values or {}),
        "python_expected_cache_mismatch_count": python_expected_cache_mismatch_count,
        "sidecar_expected_cache_mismatch_count": sidecar_expected_cache_mismatch_count,
        "expected_cache_mismatch_count": expected_cache_mismatch_count,
        "numeric_relative_tolerance": 1e-15,
        "elapsed_ms": round((time.perf_counter() - started) * 1000),
        "mismatch_count": mismatch_count,
    }


def range_stats(template_path: Path, ranges):
    workbook = load_workbook(template_path, data_only=False, read_only=False)
    formulas = 0
    comments = 0
    for item in ranges:
        min_col, min_row, max_col, max_row = range_boundaries(item["address"].replace("$", ""))
        sheet = workbook[item["sheet_name"]]
        for row in sheet.iter_rows(
            min_row=min_row,
            max_row=max_row,
            min_col=min_col,
            max_col=max_col,
        ):
            for cell in row:
                formulas += int(cell.data_type == "f")
                comments += int(cell.comment is not None)
    workbook.close()
    return {"formula_cells": formulas, "range_comment_count_python": comments}


def package_difference_counts(left_path: Path, right_path: Path):
    def category(name: str) -> str:
        if name == "[Content_Types].xml":
            return "content_types"
        if name.startswith("xl/worksheets/"):
            return "worksheets"
        if name.startswith("xl/comments"):
            return "comments"
        if name.startswith("xl/drawings/"):
            return "drawings"
        if name.startswith("xl/_rels/") or name.startswith("_rels/"):
            return "relationships"
        if name == "xl/styles.xml":
            return "styles"
        if name == "xl/workbook.xml":
            return "workbook"
        if name.startswith("docprops/"):
            return "properties"
        return "other"

    with zipfile.ZipFile(left_path) as left, zipfile.ZipFile(right_path) as right:
        left_names = set(left.namelist())
        right_names = set(right.namelist())
        common = left_names & right_names
        name_difference_counts = {}
        byte_difference_counts = {}
        for name in left_names ^ right_names:
            bucket = category(name)
            name_difference_counts[bucket] = name_difference_counts.get(bucket, 0) + 1
        for name in common:
            if left.read(name) != right.read(name):
                bucket = category(name)
                byte_difference_counts[bucket] = byte_difference_counts.get(bucket, 0) + 1
        return {
            "left_part_count": len(left_names),
            "right_part_count": len(right_names),
            "part_name_difference_count": len(left_names ^ right_names),
            "part_name_difference_category_counts": dict(sorted(name_difference_counts.items())),
            "common_part_byte_difference_count": sum(byte_difference_counts.values()),
            "common_part_byte_difference_category_counts": dict(sorted(byte_difference_counts.items())),
        }


def invoke_candidate(workspace: Path, paths: dict[str, Path], distribution: str):
    if os.name == "nt":
        assignments = " ".join(
            f"{key}={shlex.quote(wsl_path(value))}" for key, value in paths.items()
        )
        command = (
            f"cd {shlex.quote(wsl_path(workspace))} && {assignments} "
            "cargo test -p audit-core --test s1_formula_copy_candidate --locked -- --ignored --nocapture"
        )
        completed = subprocess.run(
            ["wsl.exe", "--distribution", distribution, "--exec", "bash", "-lc", command],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            check=False,
        )
    else:
        environment = os.environ.copy()
        environment.update({key: str(value) for key, value in paths.items()})
        completed = subprocess.run(
            ["cargo", "test", "-p", "audit-core", "--test", "s1_formula_copy_candidate", "--locked", "--", "--ignored", "--nocapture"],
            cwd=workspace,
            env=environment,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            check=False,
        )
    return completed.returncode


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--source-root", type=Path, required=True)
    parser.add_argument("--template-workbook", type=Path)
    parser.add_argument("--audit-workbook", type=Path)
    parser.add_argument("--mapping-config", type=Path)
    parser.add_argument("--workspace", type=Path, default=ROOT)
    parser.add_argument("--wsl-distribution", default="Ubuntu")
    parser.add_argument("--report", type=Path, required=True)
    parser.add_argument(
        "--sidecar-path",
        type=Path,
        help="Explicit lo-convert binary; when supplied, compare Python and sidecar LibreOffice formula caches",
    )
    args = parser.parse_args()
    if bool(args.template_workbook) != bool(args.audit_workbook):
        parser.error("provide both real workbook inputs or neither")

    source_root = args.source_root.resolve(strict=True)
    sidecar_path = args.sidecar_path.resolve(strict=True) if args.sidecar_path else None
    if sidecar_path is not None and not sidecar_path.is_file():
        parser.error("--sidecar-path must point to the lo-convert executable")
    sys.path.insert(0, str(source_root / "core" / "src"))
    workspace = args.workspace.resolve(strict=True)
    report_path = args.report.resolve()

    with tempfile.TemporaryDirectory(prefix="audit-s1-formula-copy-candidate-") as temporary:
        directory = Path(temporary)
        if args.template_workbook:
            template_path = args.template_workbook.resolve(strict=True)
            audit_path = args.audit_workbook.resolve(strict=True)
            mapping_path = args.mapping_config.resolve(strict=True) if args.mapping_config else None
            before_template = template_path.read_bytes()
            before_audit = audit_path.read_bytes()
            definition = resolve_real_ranges(source_root, template_path, mapping_path)
            ranges = [
                {"sheet_name": item.sheet_name, "address": item.address}
                for item in definition.copy_ranges
            ]
            sample_kind = "real_named_formula_ranges"
        else:
            template_path, audit_path, ranges = make_synthetic_pair(directory)
            before_template = template_path.read_bytes()
            before_audit = audit_path.read_bytes()
            definition = SimpleNamespace(
                rules=[],
                copy_ranges=[SimpleNamespace(**item) for item in ranges],
                structured=False,
            )
            mapping_path = None
            sample_kind = "synthetic"

        if len(ranges) != 4:
            raise ValueError("candidate must contain exactly four ImportRange operations")
        python_output = directory / "python_output.xlsx"
        rust_output = directory / "rust_output.xlsx"
        rust_report_path = directory / "rust_candidate.json"
        shutil.copy2(audit_path, python_output)

        from base_audit.native.openpyxl_workbook import copy_formula_ranges_from_template

        template_book = load_workbook(template_path, data_only=False, keep_links=False)
        python_started = time.perf_counter()
        copy_formula_ranges_from_template(template_book, python_output, definition)
        python_elapsed_ms = round((time.perf_counter() - python_started) * 1000)
        template_book.close()

        rust_started = time.perf_counter()
        candidate_paths = {
            "AUDIT_S1_COPY_CANDIDATE_INPUT": audit_path,
            "AUDIT_S1_COPY_CANDIDATE_TEMPLATE": template_path,
            "AUDIT_S1_COPY_CANDIDATE_OUTPUT": rust_output,
            "AUDIT_S1_COPY_CANDIDATE_REPORT": rust_report_path,
        }
        if mapping_path is not None:
            candidate_paths["AUDIT_S1_COPY_CANDIDATE_MAPPING"] = mapping_path
        rust_exit = invoke_candidate(workspace, candidate_paths, args.wsl_distribution)
        rust_elapsed_ms = round((time.perf_counter() - rust_started) * 1000)
        candidate = json.loads(rust_report_path.read_text(encoding="utf-8")) if rust_report_path.exists() else {
            "status": "candidate_invocation_failed",
            "production_route_eligible": False,
        }

        recalculation = {"status": "not_requested", "mismatch_count": 0}
        if sidecar_path is not None:
            expected_cache_values = (
                {
                    ("审核表", "A1"): 3,
                    ("审核表", "D1"): 7,
                    ("审核表", "G1"): 11,
                    ("审核表", "J1"): 15,
                }
                if sample_kind == "synthetic"
                else None
            )
            if rust_output.exists():
                try:
                    recalculation = recalculate_and_compare(
                        sidecar_path,
                        python_output,
                        rust_output,
                        directory,
                        expected_cache_values,
                    )
                except CandidateReadbackError as error:
                    recalculation = {
                        "status": "failed",
                        "error_category": str(error),
                        "mismatch_count": 1,
                    }
            else:
                recalculation = {
                    "status": "failed",
                    "error_category": "rust_output_missing",
                    "mismatch_count": 1,
                }

        readback_failed = False
        readback_failure_category = None
        if rust_output.exists():
            python_calc = calculation_flags(python_output)
            python_cache_count = formula_cache_count(python_output)
            python_style_references = style_reference_summary(python_output)
            try:
                rust_calc = calculation_flags(rust_output)
                package_differences = package_difference_counts(python_output, rust_output)
                candidate_cache_count = formula_cache_count(rust_output)
                rust_style_references = style_reference_summary(rust_output)
                calc_mismatch = int(python_calc != rust_calc)
            except Exception as error:
                readback_failed = True
                readback_failure_category = f"package_{type(error).__name__}"
                rust_calc = {"full_calc_on_load": None, "force_full_calc": None}
                package_differences = {
                    "left_part_count": 0,
                    "right_part_count": 0,
                    "part_name_difference_count": 0,
                    "common_part_byte_difference_count": 0,
                }
                candidate_cache_count = 0
                rust_style_references = {
                    "cell_xf_count": 0,
                    "cell_style_reference_count": 0,
                    "invalid_cell_style_reference_count": 0,
                    "max_cell_style_index": -1,
                    "sheets_with_invalid_style_references": 0,
                    "invalid_style_component_reference_counts": {},
                }
                calc_mismatch = 1
            try:
                cell_comparison = cell_mismatch_counts(python_output, rust_output, ranges)
            except CandidateReadbackError as error:
                readback_failed = True
                readback_failure_category = readback_failure_category or str(error)
                cell_comparison = {
                    "range_cells": 0, "range_mismatches": 0,
                    "outside_cells": 0, "outside_mismatches": 0,
                }
            try:
                python_metadata = workbook_metadata(python_output)
                rust_metadata = workbook_metadata(rust_output)
                metadata_mismatch = int(python_metadata != rust_metadata)
            except Exception as error:
                readback_failed = True
                readback_failure_category = readback_failure_category or f"metadata_{type(error).__name__}"
                metadata_mismatch = 1
        else:
            readback_failed = True
            readback_failure_category = "rust_output_missing"
            cell_comparison = {
                "range_cells": 0, "range_mismatches": 0,
                "outside_cells": 0, "outside_mismatches": 0,
            }
            metadata_mismatch = 1
            calc_mismatch = 1
            python_calc = calculation_flags(python_output)
            rust_calc = {"full_calc_on_load": None, "force_full_calc": None}
            package_differences = {"left_part_count": 0, "right_part_count": 0, "part_name_difference_count": 0, "common_part_byte_difference_count": 0}
            candidate_cache_count = 0
            python_cache_count = formula_cache_count(python_output)
            python_style_references = style_reference_summary(python_output)
            rust_style_references = {
                "cell_xf_count": 0,
                "cell_style_reference_count": 0,
                "invalid_cell_style_reference_count": 0,
                "max_cell_style_index": -1,
                "sheets_with_invalid_style_references": 0,
                "invalid_style_component_reference_counts": {},
            }

        inputs_unchanged = (
            template_path.read_bytes() == before_template
            and audit_path.read_bytes() == before_audit
        )
        mismatches = (
            cell_comparison["range_mismatches"]
            + cell_comparison["outside_mismatches"]
            + metadata_mismatch
            + calc_mismatch
            + int(readback_failed)
        )
        invalid_style_reference_count = (
            python_style_references["invalid_cell_style_reference_count"]
            + rust_style_references["invalid_cell_style_reference_count"]
        )
        mismatches += invalid_style_reference_count + recalculation["mismatch_count"]
        scope_stats = range_stats(template_path, ranges)
        report = {
            "schema_version": 1,
            "suite": "audit.s1_formula_copy_candidate",
            "status": "candidate_match_unverified" if rust_exit == 0 and mismatches == 0 and inputs_unchanged else "candidate_mismatch_or_unavailable",
            "sample_kind": sample_kind,
            "python": {"baseline": "base_audit.native.openpyxl_workbook.copy_formula_ranges_from_template", "elapsed_ms": python_elapsed_ms},
            "rust_candidate": {
                **candidate,
                "elapsed_ms_including_test_harness": rust_elapsed_ms,
                "candidate_test_exit_code": rust_exit,
                "formula_cache_nonempty_count": candidate_cache_count,
                "output_readback_failed": readback_failed,
                "output_readback_failure_category": readback_failure_category if readback_failed else None,
            },
            "copy_scope": {
                "range_count": len(ranges),
                **scope_stats,
            },
            "comparison": {
                **cell_comparison,
                "workbook_metadata_mismatch_count": metadata_mismatch,
                "calculation_flags_mismatch_count": calc_mismatch,
                "python_calculation_flags": python_calc,
                "rust_calculation_flags": rust_calc,
                "python_formula_cache_nonempty_count": python_cache_count,
                "package_differences_python_vs_rust": package_differences,
                "python_style_references": python_style_references,
                "rust_style_references": rust_style_references,
                "invalid_style_reference_count": invalid_style_reference_count,
                "output_readback_failed": readback_failed,
                "mismatch_count": mismatches,
            },
            "recalculation": recalculation,
            "verification": {
                "input_workbooks_unchanged": inputs_unchanged,
                "formula_recalculation_performed": recalculation["status"] in ("matched", "mismatch"),
                "capability_status": "unverified",
                "production_route_eligible": False,
                "target_linux_validation": "not performed",
                "scope": "four cross-workbook ImportRange operations plus calculation flags; direct candidate only",
            },
        }

    report_path.parent.mkdir(parents=True, exist_ok=True)
    report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(
        json.dumps(
            {
                "status": report["status"],
                "mismatch_count": mismatches,
                "sample_kind": sample_kind,
                "recalculation_status": recalculation["status"],
            }
        )
    )
    return 0 if inputs_unchanged else 1


if __name__ == "__main__":
    raise SystemExit(main())
