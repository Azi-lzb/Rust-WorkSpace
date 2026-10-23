"""Synthetic Python/Rust parity probe for pytools summary tools 6-1 and 6-2."""

from __future__ import annotations

import contextlib
import datetime as dt
import json
import logging
import math
import shlex
import subprocess
import sys
import tempfile
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.comments import Comment

sys.dont_write_bytecode = True
WORKSPACE = Path(__file__).resolve().parents[2]
PYTOOLS_SOURCE = Path(r"C:\Users\AZI\Desktop\VibeCode\4.pytools-main")


def wsl_path(path: Path) -> str:
    return f"/mnt/{path.drive.rstrip(':').lower()}/{path.as_posix()[2:].lstrip('/')}"


def from_wsl_path(value: str) -> Path:
    normalized = value.replace("\\", "/")
    if not normalized.startswith("/mnt/"):
        raise ValueError(f"Rust output outside WSL mounted drive: {value}")
    return Path(f"{normalized[5].upper()}:\\{normalized[7:].replace('/', chr(92))}")


def from_wsl_path(value: str) -> Path:
    normalized = value.replace("\\", "/")
    if not normalized.startswith("/mnt/"):
        raise ValueError(f"Rust output outside WSL mounted drive: {value}")
    return Path(f"{normalized[5].upper()}:\\{normalized[7:].replace('/', chr(92))}")


def make_source(path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Branch_A"
    ws["A1"] = "机构"
    ws["B1"] = "余额"
    ws["C1"] = "余额"
    ws["D1"] = "华东"
    ws["A2"] = "甲"
    ws["B2"] = 10.5
    ws["C2"] = 20
    ws["A3"] = "乙"
    ws["B3"] = dt.date(2026, 6, 30)
    ws["C3"] = True
    ws["A4"] = "丙"
    ws["B4"] = "text"
    ws["C4"] = 0
    ws["H10"] = "dimension tail"
    ws["H10"].number_format = "General"

    matching = wb.create_sheet("模板")
    matching["A1"] = "sheet"
    matching["B2"] = 5
    matching["C2"] = 6
    matching["D1"] = "华南"
    wb.save(path)


def make_comment_template(path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "模板"
    ws["A1"] = "单位"
    ws["A2"] = "区域起"
    ws["A4"] = "区域止"
    ws["A2"].comment = Comment("行区域1", "probe")
    ws["A4"].comment = Comment("行区域#1", "probe")
    ws["B1"] = "余额"
    ws["C1"] = "余额"
    ws["B1"].comment = Comment("列区域1", "probe")
    ws["C1"].comment = Comment("列区域#1", "probe")
    ws["D1"] = "区域"
    ws["D1"].comment = Comment("set(区域)", "probe")
    wb.save(path)


def python_baselines(source: Path, template: Path, output: Path) -> tuple[Path | None, Path | None]:
    sys.path.insert(0, str(PYTOOLS_SOURCE / "PyCode"))
    from core.summary_321_322 import run_summary_by_comment, run_summary_by_usedrange

    output.mkdir(parents=True, exist_ok=True)
    with contextlib.redirect_stdout(None):
        used = run_summary_by_usedrange([source], output / "python_used", output / "python_logs")
        comments = run_summary_by_comment(template, [source], output / "python_comments", output / "python_logs")
    logging.shutdown()
    return (
        Path(used["path"]) if used.get("saved") else None,
        Path(comments["path"]) if comments.get("saved") else None,
    )


def run_rust(command_name: str, args: list[Path], output: Path) -> tuple[Path | None, dict[str, object]]:
    cli = " ".join(shlex.quote(wsl_path(path)) for path in args)
    command = (
        f"cd {shlex.quote(wsl_path(WORKSPACE))} && cargo run --quiet -p pytools-core --bin pytools -- "
        f"{command_name} --output-dir {shlex.quote(wsl_path(output))} {cli}"
    )
    completed = subprocess.run(["wsl", "-e", "sh", "-lc", command], check=True, capture_output=True)
    stdout = completed.stdout.decode("utf-8-sig", errors="replace").replace("\x00", "")
    stderr = completed.stderr.decode("utf-8-sig", errors="replace").replace("\x00", "")
    for index, character in enumerate(stdout):
        if character != "{":
            continue
        try:
            result = json.loads(stdout[index:])
            path = result.get("path")
            return (from_wsl_path(path) if path else None), result
        except json.JSONDecodeError:
            continue
    raise RuntimeError(f"Rust summary CLI emitted no JSON: {stdout}\n{stderr}")


def run_rust_comments(template: Path, source: Path, output: Path) -> tuple[Path | None, dict[str, object]]:
    command = (
        f"cd {shlex.quote(wsl_path(WORKSPACE))} && cargo run --quiet -p pytools-core --bin pytools -- "
        f"summary-comments --template {shlex.quote(wsl_path(template))} "
        f"--output-dir {shlex.quote(wsl_path(output))} {shlex.quote(wsl_path(source))}"
    )
    completed = subprocess.run(["wsl", "-e", "sh", "-lc", command], check=True, capture_output=True)
    stdout = completed.stdout.decode("utf-8-sig", errors="replace").replace("\x00", "")
    stderr = completed.stderr.decode("utf-8-sig", errors="replace").replace("\x00", "")
    for index, character in enumerate(stdout):
        if character != "{":
            continue
        try:
            result = json.loads(stdout[index:])
            path = result.get("path")
            return (from_wsl_path(path) if path else None), result
        except json.JSONDecodeError:
            continue
    raise RuntimeError(f"Rust comment-summary CLI emitted no JSON: {stdout}\n{stderr}")


def normalize(value: object) -> object:
    if isinstance(value, (dt.datetime, dt.date, dt.time)):
        return value.isoformat()
    if isinstance(value, float) and not math.isfinite(value):
        return str(value)
    return value


def workbook_projection(path: Path) -> tuple[list[dict[str, object]], list[dict[str, object]]]:
    wb = load_workbook(path, data_only=False)
    sheets = []
    for ws in wb.worksheets:
        rows = [[normalize(value) for value in row] for row in ws.iter_rows(values_only=True)]
        sheets.append({
            "name": ws.title,
            "rows": rows,
            "freeze_panes": ws.freeze_panes,
            "header_bold": [ws.cell(1, column).font.bold for column in range(1, ws.max_column + 1)],
        })
    wb.close()
    return sheets, [
        {"name": sheet["name"], "rows": len(sheet["rows"]), "columns": max((len(row) for row in sheet["rows"]), default=0)}
        for sheet in sheets
    ]


def assert_equivalent(expected: Path | None, actual: Path | None, label: str) -> None:
    if expected is None or actual is None:
        if expected is not actual:
            raise AssertionError(f"{label}: saved status differs Python={expected}, Rust={actual}")
        return
    expected_sheets, expected_dimensions = workbook_projection(expected)
    actual_sheets, actual_dimensions = workbook_projection(actual)
    if expected_sheets != actual_sheets or expected_dimensions != actual_dimensions:
        raise AssertionError(json.dumps({
            "label": label,
            "python": expected_sheets,
            "rust": actual_sheets,
        }, ensure_ascii=False, default=str, indent=2))


def main() -> int:
    if not PYTOOLS_SOURCE.is_dir():
        raise FileNotFoundError(PYTOOLS_SOURCE)
    with tempfile.TemporaryDirectory(prefix="pytools-summary-parity-", ignore_cleanup_errors=True) as temp:
        root = Path(temp)
        source = root / "sample_2026-06-30.xlsx"
        template = root / "comment_template.xlsx"
        make_source(source)
        make_comment_template(template)
        python_used, python_comments = python_baselines(source, template, root)
        rust_used, used_report = run_rust("summary-used-range", [source], root / "rust_used")
        rust_comments, comments_report = run_rust_comments(template, source, root / "rust_comments")
        assert_equivalent(python_used, rust_used, "6-1 used-range summary")
        assert_equivalent(python_comments, rust_comments, "6-2 comment summary")
        print(json.dumps({
            "status": "passed",
            "used_range": {"python": python_used is not None, "rust": rust_used is not None},
            "comment_summary": {"python": python_comments is not None, "rust": rust_comments is not None},
            "used_range_rows_rust": used_report.get("rows"),
            "comment_summary_rows_rust": comments_report.get("rows"),
        }, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
