#!/usr/bin/env python3
"""Compare the Rust S1-F02 named-region CLI with the read-only Python flow."""

from __future__ import annotations

import argparse
import hashlib
import json
import subprocess
import sys
import tempfile
import time
from datetime import date, datetime
from pathlib import Path
from typing import Any

sys.dont_write_bytecode = True
WORKSPACE = Path(__file__).resolve().parents[2]
SOURCE_CORE = Path(r"C:\Users\AZI\Desktop\VibeCode\3.open-data-audit\core")
sys.path.insert(0, str(SOURCE_CORE / "src"))

from base_audit.name_config import (  # noqa: E402
    FIXED_ROW_SUMMARY_FUNCTION,
    USED_RANGE_SUMMARY_FUNCTION,
    FeatureMapping,
)
from base_audit.region_summary import run_region_summaries  # noqa: E402
from base_audit.excel_com import ExcelSession  # noqa: E402
from openpyxl import Workbook, load_workbook  # noqa: E402
from openpyxl.workbook.defined_name import DefinedName  # noqa: E402


def _open_workbook_for_parity(
    session: ExcelSession, path: Path, *, read_only: bool, purpose: str = ""
) -> Any:
    """Use Excel's supported named arguments for this Windows parity run.

    This Office installation rejects the source adapter's named `UpdateLinks`
    argument even though it opens the same workbook when that option is left
    at its default and `AskToUpdateLinks` is disabled on the application.
    The adapter changes only the COM call shape, not the Python summary logic.
    """
    opened_at = time.monotonic()
    workbook = session.excel.Workbooks.Open(
        str(path.resolve()),
        ReadOnly=read_only,
        IgnoreReadOnlyRecommended=True,
    )
    elapsed = time.monotonic() - opened_at
    if purpose == "template":
        session.lifecycle_metrics["template_com_open_time"] += elapsed
    if not session._first_workbook_opened:
        session._first_workbook_opened = True
        if session._ready_at is not None:
            session.lifecycle_metrics["idle_wait_time"] = max(
                0.0, opened_at - session._ready_at
            )
    return workbook


ExcelSession.open_workbook = _open_workbook_for_parity


def digest(path: Path) -> str:
    result = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            result.update(chunk)
    return result.hexdigest()


def write_source(
    path: Path,
    rows: dict[int, tuple[Any, Any]],
    local_value: str,
    global_value: str | None = None,
) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Data"
    sheet["F1"] = local_value
    for row_number, values in rows.items():
        sheet.cell(row_number, 2, values[0])
        sheet.cell(row_number, 3, values[1])
    if global_value is not None:
        metadata_sheet = workbook.create_sheet("Meta")
        metadata_sheet["A1"] = global_value
    workbook.save(path)


def equal_value(left: Any, right: Any) -> bool:
    if isinstance(left, (date, datetime)):
        left = left.isoformat()
    if isinstance(right, (date, datetime)):
        right = right.isoformat()
    if isinstance(left, bool) or isinstance(right, bool):
        return left is right
    if isinstance(left, (int, float)) and isinstance(right, (int, float)):
        return left == right or abs(left - right) <= max(abs(left), abs(right), 1.0) * 1e-15
    return left == right


def _color_signature(color: Any) -> tuple[Any, ...] | None:
    if color is None:
        return None
    color_type = color.type
    value = (
        color.rgb
        if color_type == "rgb"
        else color.indexed
        if color_type == "indexed"
        else color.theme
        if color_type == "theme"
        else None
    )
    return color_type, str(value), float(color.tint or 0)


def _font_signature(font: Any) -> tuple[Any, ...]:
    return (
        font.name,
        font.sz,
        font.bold,
        font.italic,
        font.underline,
        font.strike,
        font.vertAlign,
        _color_signature(font.color),
        font.charset,
        font.family,
    )


def _alignment_signature(alignment: Any) -> tuple[Any, ...]:
    canonical = lambda value: None if value == "general" else value
    return (
        canonical(alignment.horizontal),
        canonical(alignment.vertical),
        alignment.textRotation,
        alignment.wrapText,
        alignment.shrinkToFit,
        alignment.indent,
        alignment.relativeIndent,
        alignment.justifyLastLine,
        alignment.readingOrder,
    )


def _column_width(sheet: Any, column: int) -> float:
    """Resolve OOXML column spans instead of treating a shared span as column A only."""
    for dimension in sheet.column_dimensions.values():
        if dimension.min is not None and dimension.max is not None:
            if dimension.min <= column <= dimension.max:
                return float(dimension.width)
    default_width = sheet.sheet_format.defaultColWidth
    return float(default_width) if default_width is not None else 13.0


def compare_workbooks(python_path: Path, rust_path: Path) -> dict[str, Any]:
    python = load_workbook(python_path, data_only=False, read_only=False)
    rust = load_workbook(rust_path, data_only=False, read_only=False)
    try:
        sheet_name_mismatches = python.sheetnames != rust.sheetnames
        value_mismatches = 0
        format_mismatches = 0
        font_mismatches = 0
        alignment_mismatches = 0
        alignment_header_mismatches = 0
        alignment_body_mismatches = 0
        alignment_property_mismatches = {
            "horizontal": 0,
            "vertical": 0,
            "textRotation": 0,
            "wrapText": 0,
            "shrinkToFit": 0,
            "indent": 0,
            "relativeIndent": 0,
            "justifyLastLine": 0,
            "readingOrder": 0,
        }
        alignment_type_pairs: dict[str, int] = {}
        alignment_horizontal_pairs: dict[str, int] = {}
        header_style_mismatches = 0
        filter_mismatches = 0
        freeze_mismatches = 0
        dimension_mismatches = 0
        width_deltas: list[float] = []
        for name in python.sheetnames:
            if name not in rust.sheetnames:
                continue
            expected, actual = python[name], rust[name]
            if (expected.max_row, expected.max_column) != (actual.max_row, actual.max_column):
                dimension_mismatches += 1
            rows = max(expected.max_row, actual.max_row)
            columns = max(expected.max_column, actual.max_column)
            for row in range(1, rows + 1):
                for column in range(1, columns + 1):
                    left, right = expected.cell(row, column), actual.cell(row, column)
                    if not equal_value(left.value, right.value):
                        value_mismatches += 1
                    if left.number_format != right.number_format:
                        format_mismatches += 1
                    if _font_signature(left.font) != _font_signature(right.font):
                        font_mismatches += 1
                    if _alignment_signature(left.alignment) != _alignment_signature(right.alignment):
                        alignment_mismatches += 1
                        if row == 1:
                            alignment_header_mismatches += 1
                        else:
                            alignment_body_mismatches += 1
                        alignment_type = f"{left.data_type}->{right.data_type}"
                        alignment_type_pairs[alignment_type] = (
                            alignment_type_pairs.get(alignment_type, 0) + 1
                        )
                        horizontal_pair = (
                            f"{left.data_type}:{left.alignment.horizontal!r}"
                            f"->{right.alignment.horizontal!r}"
                        )
                        alignment_horizontal_pairs[horizontal_pair] = (
                            alignment_horizontal_pairs.get(horizontal_pair, 0) + 1
                        )
                        for property_name in alignment_property_mismatches:
                            if getattr(left.alignment, property_name) != getattr(
                                right.alignment, property_name
                            ):
                                alignment_property_mismatches[property_name] += 1
                    if row == 1:
                        left_color = str(left.fill.fgColor.rgb)[-6:]
                        right_color = str(right.fill.fgColor.rgb)[-6:]
                        if (
                            left.font.bold != right.font.bold
                            or left.fill.fill_type != right.fill.fill_type
                            or left_color != right_color
                        ):
                            header_style_mismatches += 1
            if expected.auto_filter.ref != actual.auto_filter.ref:
                filter_mismatches += 1
            if expected.freeze_panes != actual.freeze_panes:
                freeze_mismatches += 1
            for column in range(1, columns + 1):
                left = _column_width(expected, column)
                right = _column_width(actual, column)
                width_deltas.append(abs(left - right))
        return {
            "status": "passed"
            if not any(
                (
                    sheet_name_mismatches,
                    value_mismatches,
                    format_mismatches,
                    font_mismatches,
                    alignment_mismatches,
                    header_style_mismatches,
                    filter_mismatches,
                    freeze_mismatches,
                    dimension_mismatches,
                )
            )
            else "failed",
            "sheet_name_mismatches": int(sheet_name_mismatches),
            "python_sheet_names": python.sheetnames,
            "rust_sheet_names": rust.sheetnames,
            "dimension_mismatches": dimension_mismatches,
            "value_mismatches": value_mismatches,
            "number_format_mismatches": format_mismatches,
            "font_mismatches": font_mismatches,
            "alignment_mismatches": alignment_mismatches,
            "alignment_header_mismatches": alignment_header_mismatches,
            "alignment_body_mismatches": alignment_body_mismatches,
            "alignment_property_mismatches": alignment_property_mismatches,
            "alignment_type_pairs": alignment_type_pairs,
            "alignment_horizontal_pairs": alignment_horizontal_pairs,
            "header_style_mismatches": header_style_mismatches,
            "filter_mismatches": filter_mismatches,
            "freeze_mismatches": freeze_mismatches,
            "column_width_max_delta": max(width_deltas, default=0.0),
            "column_widths_exact": all(delta == 0 for delta in width_deltas),
            "data_rows_by_sheet": {
                name: max(0, python[name].max_row - 1) for name in python.sheetnames
            },
        }
    finally:
        python.close()
        rust.close()


def write_history_case(root: Path) -> tuple[Path, Path, Path, Path]:
    template_path = root / "history_template.xlsx"
    template = Workbook()
    template_sheet = template.active
    template_sheet.title = "本地校验结果"
    for column, header in enumerate(
        (
            "批次",
            "表单名称",
            "规则编号",
            "规则类型",
            "规则描述",
            "校验字段",
        ),
        start=2,
    ):
        template_sheet.cell(2, column, header)
    template.defined_names.add(
        DefinedName("任意行汇总区域", attr_text="'本地校验结果'!$B$3:$G$10")
    )
    template.defined_names.add(
        DefinedName("表头区域", attr_text="'本地校验结果'!$B$2:$G$2")
    )
    template.save(template_path)
    template.close()

    input_dir = root / "history_input"
    input_dir.mkdir()
    source_path = input_dir / "Sample.xlsx"
    source = Workbook()
    source_sheet = source.active
    source_sheet.title = "本地校验结果"
    source_rows = (
        ("2026-07", "存量单位贷款", 1.0, "阈值", "不应低于", "贷款余额"),
        ("2026-07", "存量单位贷款", 2.0, "阈值", "不应低于", "贷款余额"),
    )
    for row_number, values in enumerate(source_rows, start=3):
        for column, value in enumerate(values, start=2):
            source_sheet.cell(row_number, column, value)
    source.save(source_path)
    source.close()

    history_path = root / "history.xlsx"
    history = Workbook()
    history_sheet = history.active
    history_sheet.title = "本地校验结果"
    history_sheet.append(
        (
            "来源文件",
            "来源工作表",
            "批次",
            "表单名称",
            "规则编号",
            "规则类型",
            "规则描述",
            "校验字段",
            "历史触发条数",
            "历史分类说明",
            "审核结果",
        )
    )
    identity = (
        "Sample",
        "本地校验结果",
        "2026-07",
        "存量单位贷款",
        "1",
        "阈值",
        "不应低于",
        "贷款余额",
    )
    history_sheet.append((*identity, 10, "类别甲", "通过"))
    history_sheet.append((*identity, 12, "类别乙", "未通过"))
    history.save(history_path)
    history.close()
    return template_path, input_dir, history_path, source_path


def write_fallback_case(root: Path) -> tuple[Path, Path, Path, Path]:
    template_path = root / "fallback_template.xlsx"
    template = Workbook()
    fallback = template.active
    fallback.title = "兜底模板"
    fallback["B1"] = "字段甲"
    fallback["C1"] = "字段乙"
    template.create_sheet("Meta")
    template.defined_names.add(
        DefinedName("任意行汇总区域", attr_text="'兜底模板'!$B$2:$C$100")
    )
    template.defined_names.add(
        DefinedName("固定行汇总区域", attr_text="'兜底模板'!$B$2:$C$6")
    )
    template.defined_names.add(
        DefinedName("表头区域", attr_text="'兜底模板'!$B$1:$C$1")
    )
    template.defined_names.add(DefinedName("全局单元格区域.批次", attr_text="Meta!$A$1"))
    template.defined_names.add(
        DefinedName("单元格区域.区域", attr_text="'兜底模板'!$F$1")
    )
    template.save(template_path)
    template.close()

    input_dir = root / "fallback_input"
    input_dir.mkdir()
    source_path = input_dir / "FallbackSource.xlsx"
    source = Workbook()
    unmatched = source.active
    unmatched.title = "未匹配表"
    unmatched["F1"] = "区域甲"
    unmatched["B4"] = "记录甲"
    unmatched["C4"] = 11
    unmatched["B6"] = "记录乙"
    unmatched["C6"] = 0
    empty = source.create_sheet("无兜底数据")
    empty["A1"] = "不在兜底列范围内"
    other = source.create_sheet("另一张表")
    other["F1"] = "区域乙"
    other["B2"] = "另一条记录"
    other["C2"] = 12
    metadata = source.create_sheet("Meta")
    metadata["A1"] = "批次-01"
    source.save(source_path)
    source.close()

    second_source_path = input_dir / "FallbackSource2.xlsx"
    second_source = Workbook()
    repeated = second_source.active
    repeated.title = "未匹配表"
    repeated["F1"] = "区域丙"
    repeated["B2"] = "记录丙"
    repeated["C2"] = 5
    second_metadata = second_source.create_sheet("Meta")
    second_metadata["A1"] = "批次-02"
    second_source.save(second_source_path)
    second_source.close()
    return template_path, input_dir, source_path, second_source_path


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--audit-binary", type=Path, required=True)
    parser.add_argument("--report", type=Path)
    args = parser.parse_args()
    binary = args.audit_binary.resolve(strict=True)
    template_path = WORKSPACE / "tests" / "fixtures" / "s1_region_summary_template.xlsx"

    with tempfile.TemporaryDirectory(prefix="audit-s1-region-summary-") as temporary:
        root = Path(temporary)
        input_dir = root / "input"
        output_dir = root / "outputs"
        input_dir.mkdir()
        output_dir.mkdir()
        alpha = input_dir / "!Alpha_20260731.xlsx"
        beta = input_dir / "Beta_20260731.xlsx"
        write_source(
            alpha,
            {4: ("甲", 10), 5: (None, 0), 7: ("乙", 20)},
            "甲组",
            "批次甲",
        )
        write_source(beta, {4: ("丙", 30), 5: ("丁", 40)}, "乙组")
        input_files = [alpha, beta]
        before = {path: digest(path) for path in input_files}
        features = (
            FeatureMapping(
                "任意行汇总",
                USED_RANGE_SUMMARY_FUNCTION,
                ("任意行汇总区域",),
                False,
                "",
                "",
            ),
            FeatureMapping(
                "固定行汇总",
                FIXED_ROW_SUMMARY_FUNCTION,
                ("固定行汇总区域",),
                False,
                "",
                "",
            ),
        )

        progress: list[str] = []
        python_started = time.perf_counter()
        result = run_region_summaries(
            template_path=template_path,
            input_dir=input_dir,
            output_dir=output_dir / "python",
            config_path=template_path,
            summary_features=features,
            on_step=progress.append,
            engine_preference="自动",
        )
        python_elapsed_ms = round((time.perf_counter() - python_started) * 1000)

        rust_path = output_dir / "rust.xlsx"
        rust_started = time.perf_counter()
        command = [
            str(binary),
            "s1-region-summary",
            "--template-workbook",
            str(template_path),
            "--input-dir",
            str(input_dir),
            "--output-workbook",
            str(rust_path),
            "--feature-name",
            "任意行汇总",
            "--feature-name",
            "固定行汇总",
            "--developer-view",
        ]
        completed = subprocess.run(
            command,
            check=False,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            encoding="utf-8",
        )
        rust_elapsed_ms = round((time.perf_counter() - rust_started) * 1000)
        if completed.returncode:
            raise RuntimeError(f"Rust CLI failed: {completed.stderr.strip()}")
        rust_payload = json.loads(completed.stdout)
        rust_events = [
            json.loads(line)
            for line in completed.stderr.splitlines()
            if line.strip().startswith("{")
        ]

        comparison = compare_workbooks(result.output_path, rust_path)

        history_root = root / "history_case"
        history_root.mkdir()
        (
            history_template,
            history_input_dir,
            history_path,
            history_source,
        ) = write_history_case(history_root)
        history_inputs = [history_template, history_path, history_source]
        history_before = {path: digest(path) for path in history_inputs}
        history_feature = FeatureMapping(
            "任意行汇总",
            USED_RANGE_SUMMARY_FUNCTION,
            ("任意行汇总区域",),
            False,
            "",
            "",
        )
        history_progress: list[str] = []
        history_python = run_region_summaries(
            template_path=history_template,
            input_dir=history_input_dir,
            output_dir=history_root / "python",
            config_path=history_template,
            summary_features=(history_feature,),
            history_config_path=history_path,
            on_step=history_progress.append,
            engine_preference="自动",
        )
        history_rust_path = history_root / "rust.xlsx"
        history_command = [
            str(binary),
            "s1-region-summary",
            "--template-workbook",
            str(history_template),
            "--input-dir",
            str(history_input_dir),
            "--output-workbook",
            str(history_rust_path),
            "--history-workbook",
            str(history_path),
            "--feature-name",
            "任意行汇总",
            "--developer-view",
        ]
        history_completed = subprocess.run(
            history_command,
            check=False,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            encoding="utf-8",
        )
        if history_completed.returncode:
            raise RuntimeError(f"Rust history CLI failed: {history_completed.stderr.strip()}")
        history_payload = json.loads(history_completed.stdout)
        history_events = [
            json.loads(line)
            for line in history_completed.stderr.splitlines()
            if line.strip().startswith("{")
        ]
        history_comparison = compare_workbooks(
            history_python.output_path, history_rust_path
        )
        history_unchanged = all(
            digest(path) == value for path, value in history_before.items()
        )
        history_status = (
            "passed"
            if history_comparison["status"] == "passed"
            and history_payload["summary"]["history_rows_matched"] == 1
            and history_unchanged
            else "failed"
        )

        fallback_root = root / "fallback_case"
        fallback_root.mkdir()
        (
            fallback_template,
            fallback_input_dir,
            fallback_source,
            fallback_second_source,
        ) = write_fallback_case(fallback_root)
        fallback_before = {
            path: digest(path)
            for path in (
                fallback_template,
                fallback_source,
                fallback_second_source,
            )
        }
        fallback_feature = FeatureMapping(
            "任意行汇总",
            USED_RANGE_SUMMARY_FUNCTION,
            ("任意行汇总区域",),
            False,
            "",
            "",
        )
        fallback_progress: list[str] = []
        fallback_python_started = time.perf_counter()
        fallback_python = run_region_summaries(
            template_path=fallback_template,
            input_dir=fallback_input_dir,
            output_dir=fallback_root / "python",
            config_path=fallback_template,
            summary_features=(fallback_feature,),
            on_step=fallback_progress.append,
            engine_preference="自动",
        )
        fallback_python_elapsed_ms = round(
            (time.perf_counter() - fallback_python_started) * 1000
        )
        fallback_rust_path = fallback_root / "rust.xlsx"
        fallback_command = [
            str(binary),
            "s1-region-summary",
            "--template-workbook",
            str(fallback_template),
            "--input-dir",
            str(fallback_input_dir),
            "--output-workbook",
            str(fallback_rust_path),
            "--feature-name",
            "任意行汇总",
            "--developer-view",
        ]
        fallback_completed = subprocess.run(
            fallback_command,
            check=False,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            encoding="utf-8",
        )
        if fallback_completed.returncode:
            raise RuntimeError(
                f"Rust fallback summary CLI failed: {fallback_completed.stderr.strip()}"
            )
        fallback_payload = json.loads(fallback_completed.stdout)
        fallback_events = [
            json.loads(line)
            for line in fallback_completed.stderr.splitlines()
            if line.strip().startswith("{")
        ]
        fallback_comparison = compare_workbooks(
            fallback_python.output_path, fallback_rust_path
        )

        fixed_fallback_feature = FeatureMapping(
            "固定行汇总",
            FIXED_ROW_SUMMARY_FUNCTION,
            ("固定行汇总区域",),
            False,
            "",
            "",
        )
        fixed_fallback_progress: list[str] = []
        fixed_fallback_python_started = time.perf_counter()
        fixed_fallback_python = run_region_summaries(
            template_path=fallback_template,
            input_dir=fallback_input_dir,
            output_dir=fallback_root / "python_fixed",
            config_path=fallback_template,
            summary_features=(fixed_fallback_feature,),
            on_step=fixed_fallback_progress.append,
            engine_preference="自动",
        )
        fixed_fallback_python_elapsed_ms = round(
            (time.perf_counter() - fixed_fallback_python_started) * 1000
        )
        fixed_fallback_rust_path = fallback_root / "rust_fixed.xlsx"
        fixed_fallback_command = [
            str(binary),
            "s1-region-summary",
            "--template-workbook",
            str(fallback_template),
            "--input-dir",
            str(fallback_input_dir),
            "--output-workbook",
            str(fixed_fallback_rust_path),
            "--feature-name",
            "固定行汇总",
            "--developer-view",
        ]
        fixed_fallback_completed = subprocess.run(
            fixed_fallback_command,
            check=False,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            encoding="utf-8",
        )
        if fixed_fallback_completed.returncode:
            raise RuntimeError(
                "Rust fixed-row fallback summary CLI failed: "
                + fixed_fallback_completed.stderr.strip()
            )
        fixed_fallback_payload = json.loads(fixed_fallback_completed.stdout)
        fixed_fallback_events = [
            json.loads(line)
            for line in fixed_fallback_completed.stderr.splitlines()
            if line.strip().startswith("{")
        ]
        fixed_fallback_comparison = compare_workbooks(
            fixed_fallback_python.output_path, fixed_fallback_rust_path
        )
        fallback_unchanged = all(
            digest(path) == value for path, value in fallback_before.items()
        )
        fallback_status = (
            "passed"
            if fallback_comparison["status"] == "passed"
            and fixed_fallback_comparison["status"] == "passed"
            and fallback_payload["summary"]["output_sheets"] == 2
            and fallback_payload["summary"]["data_rows"] == 4
            and fixed_fallback_payload["summary"]["output_sheets"] == 2
            and fixed_fallback_payload["summary"]["data_rows"] == 4
            and fallback_unchanged
            else "failed"
        )

        status = (
            "passed"
            if comparison["status"] == "passed"
            and history_status == "passed"
            and fallback_status == "passed"
            else "failed"
        )
        if status == "passed" and (
            not comparison["column_widths_exact"]
            or not history_comparison["column_widths_exact"]
            or not fallback_comparison["column_widths_exact"]
            or not fixed_fallback_comparison["column_widths_exact"]
        ):
            status = "passed_with_column_width_variance"
        report = {
            "feature": "S1-F02 named-region summary",
            "status": status,
            "fixture": "synthetic",
            "python_engine": next(
                (line.split("：", 1)[1] for line in progress if line.startswith("已连接表格引擎：")),
                "not_reported",
            ),
            "python_com_call_adapter": "test-only: this Office installation rejects the source adapter's named UpdateLinks=0 argument; the baseline used ReadOnly and IgnoreReadOnlyRecommended while Excel AskToUpdateLinks remained disabled",
            "rust_reader": rust_payload["summary"]["reader"],
            "rust_writer": rust_payload["summary"]["writer"],
            "python_elapsed_ms": python_elapsed_ms,
            "rust_cli_elapsed_ms": rust_elapsed_ms,
            "rust_events": len(rust_events),
            "source_workbooks": rust_payload["summary"]["source_workbooks"],
            "source_sheets": rust_payload["summary"]["source_sheets"],
            "metadata_fields_left_blank": rust_payload["summary"][
                "metadata_fields_left_blank"
            ],
            "output_sheets": rust_payload["summary"]["output_sheets"],
            "data_rows": rust_payload["summary"]["data_rows"],
            "input_files_unchanged": all(digest(path) == value for path, value in before.items()),
            "comparison": comparison,
            "history_enrichment": {
                "status": history_status,
                "python_engine": next(
                    (
                        line.split("：", 1)[1]
                        for line in history_progress
                        if line.startswith("已连接表格引擎：")
                    ),
                    "not_reported",
                ),
                "rust_history_reader": history_payload["summary"]["history_reader"],
                "rust_history_rows_matched": history_payload["summary"]["history_rows_matched"],
                "rust_history_elapsed_ms": history_payload["summary"]["history_elapsed_ms"],
                "rust_events": len(history_events),
                "input_files_unchanged": history_unchanged,
                "comparison": history_comparison,
                "scope": "synthetic local-validation history sheet; eight-column composite key; numeric 1.0 to text '1' key normalization; duplicate history rows joined in source order; unmatched source row remains blank",
            },
            "fallback_template": {
                "status": fallback_status,
                "python_engine": next(
                    (
                        line.split("：", 1)[1]
                        for line in fallback_progress
                        if line.startswith("已连接表格引擎：")
                    ),
                    "not_reported",
                ),
                "python_elapsed_ms": fallback_python_elapsed_ms,
                "rust_reader": fallback_payload["summary"]["reader"],
                "rust_writer": fallback_payload["summary"]["writer"],
                "rust_elapsed_ms": fallback_payload["summary"]["elapsed_ms"],
                "rust_events": len(fallback_events),
                "output_sheets": fallback_payload["summary"]["output_sheets"],
                "data_rows": fallback_payload["summary"]["data_rows"],
                "input_files_unchanged": fallback_unchanged,
                "comparison": fallback_comparison,
                "fixed_rows": {
                    "python_engine": next(
                        (
                            line.split("：", 1)[1]
                            for line in fixed_fallback_progress
                            if line.startswith("已连接表格引擎：")
                        ),
                        "not_reported",
                    ),
                    "python_elapsed_ms": fixed_fallback_python_elapsed_ms,
                    "rust_reader": fixed_fallback_payload["summary"]["reader"],
                    "rust_writer": fixed_fallback_payload["summary"]["writer"],
                    "rust_elapsed_ms": fixed_fallback_payload["summary"]["elapsed_ms"],
                    "rust_events": len(fixed_fallback_events),
                    "output_sheets": fixed_fallback_payload["summary"]["output_sheets"],
                    "data_rows": fixed_fallback_payload["summary"]["data_rows"],
                    "comparison": fixed_fallback_comparison,
                    "scope": "same source workbooks using the configured fixed-row fallback range",
                },
                "scope": "synthetic unmatched source sheets across multiple workbooks; arbitrary-row and fixed-row fallback ranges; repeated sheet names merged; local fallback-sheet metadata mapped by coordinates onto each source sheet; global metadata; blank-only unmatched sheet omitted",
            },
            "scope": "Windows synthetic formula-free XLSX; named arbitrary/fixed rows; unmatched source-sheet fallback; global/local single-cell metadata fields; missing source metadata sheet leaves the global field blank; merged multi-row header; value/type, sheet order, dimensions, number format, header style, filter and freeze comparison; built-in local-validation history append comparison; no real template or Linux runtime",
        }
        encoded = json.dumps(report, ensure_ascii=False, indent=2) + "\n"
        if args.report:
            args.report.parent.mkdir(parents=True, exist_ok=True)
            args.report.write_text(encoded, encoding="utf-8")
        print(encoded, end="")
        return (
            0
            if status.startswith("passed")
            and report["input_files_unchanged"]
            and fallback_unchanged
            else 1
        )


if __name__ == "__main__":
    raise SystemExit(main())
