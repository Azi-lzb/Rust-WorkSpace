"""Compare the Rust 1-3/1-4 wide-summary CLI with the read-only PyCode baseline.

The script builds a synthetic XLSX/config under a temporary directory, calls
the original Python workflow without writing to the source tree, then compares
both generated workbooks using openpyxl.
"""

from __future__ import annotations

import json
import gc
import hashlib
import logging
import math
import os
import shlex
import shutil
import subprocess
import sys
import tempfile
import time
from datetime import date, datetime
from importlib.metadata import version
from pathlib import Path

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


def make_input(path: Path) -> None:
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
    sheet["B5"] = 10.25
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
        ("输入目录", to_wsl_path(input_dir)),
        ("输出目录", to_wsl_path(output)),
        ("日志目录", to_wsl_path(output / "logs")),
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
        "", to_wsl_path(target), "Append", "是", "集团@B2", "", "是",
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
    workspace = to_wsl_path(WORKSPACE)
    config_wsl = to_wsl_path(config)
    source_wsl = to_wsl_path(source)
    command = (
        f"cd {shlex.quote(workspace)} && cargo run --quiet -p pytools-core --bin pytools -- "
        f"{cli} --config {shlex.quote(config_wsl)} {shlex.quote(source_wsl)}"
    )
    completed = subprocess.run(
        ["wsl", "-e", "sh", "-lc", command],
        check=True,
        capture_output=True,
    )
    stdout = completed.stdout.decode("utf-8-sig", errors="replace").replace("\x00", "")
    stderr = completed.stderr.decode("utf-8-sig", errors="replace").replace("\x00", "")
    for line in reversed(stdout.splitlines()):
        try:
            result = json.loads(line)
            value = result.get("output")
            output = from_wsl_path(value) if value else None
            plans = [from_wsl_path(path) for path in result.get("target_append_plans", [])]
            return output, plans
        except json.JSONDecodeError:
            continue
    raise RuntimeError(f"Rust CLI did not return JSON: {stdout}\n{stderr}")


def execute_new_target_plan(plan: Path) -> tuple[Path, Path, dict[str, object]]:
    workspace = to_wsl_path(WORKSPACE)
    plan_wsl = to_wsl_path(plan)
    command = (
        f"cd {shlex.quote(workspace)} && cargo run --quiet -p pytools-core --bin pytools -- "
        f"dedup-append-execute --plan {shlex.quote(plan_wsl)} --commit-target"
    )
    completed = subprocess.run(
        ["wsl", "-e", "sh", "-lc", command],
        check=True,
        capture_output=True,
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
    if isinstance(result, dict) and "candidate" in result:
        candidate = from_wsl_path(result["candidate"])
        wrapper = json.loads(plan.read_text(encoding="utf-8"))
        target = from_wsl_path(wrapper["target"])
        if result.get("target_commit", {}).get("operation") != "created":
            raise AssertionError(f"new configured target was not committed: {result}")
        if target.read_bytes() != candidate.read_bytes():
            raise AssertionError("new configured target differs from validated candidate")
        return candidate, target, result
    raise RuntimeError(f"append executor did not return a committed candidate: {stdout}\n{stderr}")


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
    workspace = to_wsl_path(WORKSPACE)
    plan_wsl = to_wsl_path(plan_path)
    capability_wsl = to_wsl_path(capability_path)
    platform = "synthetic-pt-1-3-1-4-umya-test"
    command = (
        f"cd {shlex.quote(workspace)} && cargo run --quiet -p pytools-core --bin pytools -- "
        f"dedup-append-execute --plan {shlex.quote(plan_wsl)} "
        f"--capabilities {shlex.quote(capability_wsl)} --platform {shlex.quote(platform)} --commit-target"
    )
    completed = subprocess.run(
        ["wsl", "-e", "sh", "-lc", command],
        check=True,
        capture_output=True,
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
    return from_wsl_path(result["output"]), result


def normalized(path: Path) -> tuple[list[str], list[list[list[object]]], list[list[str]]]:
    workbook = load_workbook(path, data_only=False)
    names: list[str] = []
    rows_by_sheet: list[list[list[object]]] = []
    formats_by_sheet: list[list[str]] = []
    for sheet in workbook.worksheets:
        names.append(sheet.title)
        rows: list[list[object]] = []
        formats: list[str] = []
        for row in sheet.iter_rows():
            values = []
            for cell in row:
                value = cell.value
                if isinstance(value, (datetime, date)):
                    value = value.isoformat()
                elif isinstance(value, float) and not math.isfinite(value):
                    value = str(value)
                values.append(value)
                if cell.row == 2:
                    formats.append(cell.number_format)
            rows.append(values)
        rows_by_sheet.append(rows)
        formats_by_sheet.append(formats)
    workbook.close()
    return names, rows_by_sheet, formats_by_sheet


def compare_values(expected: object, actual: object, location: str) -> None:
    if isinstance(expected, (int, float)) and isinstance(actual, (int, float)):
        if not math.isclose(float(expected), float(actual), rel_tol=1e-15, abs_tol=1e-15):
            raise AssertionError(f"{location}: {expected!r} != {actual!r}")
    elif expected != actual:
        raise AssertionError(f"{location}: {expected!r} != {actual!r}")


def assert_parity(expected_path: Path, actual_path: Path) -> None:
    expected_names, expected_rows, expected_formats = normalized(expected_path)
    actual_names, actual_rows, actual_formats = normalized(actual_path)
    assert expected_names == actual_names, (expected_names, actual_names)
    assert len(expected_rows) == len(actual_rows)
    for sheet_index, (expected_sheet, actual_sheet) in enumerate(zip(expected_rows, actual_rows)):
        assert len(expected_sheet) == len(actual_sheet)
        for row_index, (expected_row, actual_row) in enumerate(zip(expected_sheet, actual_sheet)):
            assert len(expected_row) == len(actual_row), (sheet_index, row_index, expected_row, actual_row)
            for column_index, (expected, actual) in enumerate(zip(expected_row, actual_row)):
                compare_values(expected, actual, f"sheet {sheet_index} cell {row_index},{column_index}")
    assert expected_formats == actual_formats, (expected_formats, actual_formats)


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


def sha256_file(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def run_case(root: Path, row_suffix: bool) -> None:
    case = root / ("suffix" if row_suffix else "no_suffix")
    case.mkdir()
    source = case / "sample_2024-06-30.xlsx"
    make_input(source)
    python_output = case / "python_output"
    rust_output = case / "rust_output"
    config = case / "config.xlsx"
    make_config(config, rust_output, source.parent, rust_output / "target.xlsx")
    baseline, python_target = python_baseline(source, python_output, row_suffix)
    actual, plans = run_rust(config, source, row_suffix)
    if baseline is None or actual is None:
        raise AssertionError(f"wide summary did not produce output: baseline={baseline}, rust={actual}")
    assert_parity(baseline, actual)
    if len(plans) != 1:
        raise AssertionError(f"expected one target append plan, got {plans}")
    candidate, committed_target, commit_execution = execute_new_target_plan(plans[0])
    assert_parity(python_target, candidate)
    if committed_target.read_bytes() != candidate.read_bytes():
        raise AssertionError("new target commit differs from candidate")
    print(json.dumps({"row_suffix": row_suffix, "summary_status": "passed", "target_append_status": "passed", "target_candidate": str(candidate), "target_commit": commit_execution["target_commit"], "execution_summary": commit_execution["execution_summary"]}, ensure_ascii=False))


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
    assert_parity(expected, candidate)
    assert_sheet_sentinels(expected, candidate, ("Before", "After"))
    assert_replacement_drops(seed_target, expected)
    assert_replacement_drops(seed_target, candidate)
    if rust_target.read_bytes() != candidate.read_bytes():
        raise AssertionError("configured target bytes differ from the validated candidate")
    if rust_target.read_bytes() == seed_bytes:
        raise AssertionError("Rust execution left the configured target unchanged")
    if source.read_bytes() != source_before or config.read_bytes() != config_before:
        raise AssertionError("Rust/Python parity run modified the source workbook or configuration")
    candidate_names, candidate_rows, _ = normalized(candidate)
    result = {
        "row_suffix": row_suffix,
        "summary_status": "passed",
        "existing_target_replacement_status": "passed",
        "engine": execution["engine"],
        "target_commit": execution["target_commit"],
        "execution_summary": execution["execution_summary"],
        "candidate_sha256": sha256_file(candidate),
        "python_reference_sha256": sha256_file(expected),
        "candidate_artifacts_retained": False,
        "committed_target_sha256": sha256_file(rust_target),
        "candidate_matches_committed_target_bytes": rust_target.read_bytes() == candidate.read_bytes(),
        "compared_sheet_order": candidate_names,
        "compared_cell_count": sum(len(row) for sheet in candidate_rows for row in sheet),
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
        run_case(root, False)
        run_case(root, True)
        existing_cases = [
            run_existing_target_case(root, False),
            run_existing_target_case(root, True),
        ]
        report = {
            "schema_version": 1,
            "suite": "pytools.pt_1_3_1_4_existing_target_replacement_parity",
            "environment": {
                "python": sys.version.split()[0],
                "python_packages": {
                    "numpy": version("numpy"),
                    "pandas": version("pandas"),
                    "openpyxl": version("openpyxl"),
                },
                "python_baseline": str(PYTOOLS_SOURCE),
                "rust_execution": "WSL2 cargo run via pytools CLI; Windows Python executes baseline and comparisons",
                "sample_kind": "synthetic XLSX only",
            },
            "comparison_contract": {
                "worksheet_names_order_values_and_types": "exact except numeric values compared with relative and absolute tolerance 1e-15",
                "formula_text": "comparison handles formula strings, but the synthetic fixture contains no formulas; preservation is unverified",
                "data_date_number_formats": "exact",
                "existing_target_behavior": "rebuild target worksheet from existing rows plus accepted non-duplicate rows; preserve untouched sheets",
                "target_sheet_replacement": "legacy replacement/drop behavior checked for selected style/comment/merge/row/column attributes; not full style parity",
                "untouched_sheet_properties": "selected values, A1 fill/comment, merges, row height, and column width",
                "original_target_input": "configured target is replaced only after capability execution, candidate validation, SHA-256 check, atomic sibling publication, and Calamine reopen; source and config remain unchanged",
                "test_capability": "temporary per-run capability; not registered in workspace production capabilities",
            },
            "existing_target_cases": existing_cases,
            "commands": {
                "probe": "<temporary Python 3.12.10 venv> pytools-core/probe/compare_wide_summary.py",
            },
            "limitations": [
                "No sanitized real business sample was available.",
                "Temporary capability evidence is test-only and does not register a production route.",
                "Windows native Rust and 麒麟/UOS glibc 2.28 deb remain unverified.",
                "The fixture contains no formulas; the probe compares values and selected object properties only and is not a universal workbook-preservation claim.",
            ],
        }
        report_path = WORKSPACE / "pytools-core" / "probe" / "reports" / "wide_summary_existing_target_parity_20261023.json"
        report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        print(json.dumps({"report": str(report_path), "cases": len(existing_cases)}, ensure_ascii=False))
        cleanup_probe_directory(root)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
