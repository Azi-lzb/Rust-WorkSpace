"""Synthetic Python/Rust parity probe for pytools timeline operation 1-1."""

from __future__ import annotations

import json
import logging
import shlex
import subprocess
import sys
import tempfile
from pathlib import Path

from openpyxl import Workbook, load_workbook

sys.dont_write_bytecode = True
WORKSPACE = Path(__file__).resolve().parents[2]
PYTOOLS_SOURCE = Path(r"C:\Users\AZI\Desktop\VibeCode\4.pytools-main")


def to_wsl_path(path: Path) -> str:
    drive = path.drive.rstrip(":").lower()
    if not drive:
        raise ValueError(f"expected absolute Windows path: {path}")
    return f"/mnt/{drive}/{path.as_posix()[2:].lstrip('/')}"


def from_wsl_path(value: str) -> Path:
    value = value.replace("\\", "/")
    if not value.startswith("/mnt/"):
        raise ValueError(f"expected path under WSL mounted drive: {value}")
    return Path(f"{value[5].upper()}:\\{value[7:].replace('/', chr(92))}")


def make_book(path: Path, changed: bool = False) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Report"
    ws["A2"] = "机构"
    ws["B2"] = "存款"
    ws["C2"] = "贷款"
    ws["B3"] = "定期" if changed else "活期"
    ws["C3"] = "公司"
    ws["A5"] = "甲支行"
    ws["B5"] = 11
    ws["C5"] = 21
    ws["A6"] = "丙支行" if changed else "乙支行"
    ws["B6"] = 12
    ws["C6"] = 22
    wb.save(path)


def make_config(path: Path, output: Path, inputs: Path) -> None:
    wb = Workbook()
    global_sheet = wb.active
    global_sheet.title = "全局配置"
    global_sheet.append(["配置项", "值"])
    for key, value in (
        ("输入目录", to_wsl_path(inputs)),
        ("输出目录", to_wsl_path(output)),
        ("日志目录", to_wsl_path(output / "logs")),
        ("错误策略", "continue"),
        ("源文件扩展名", ".xlsx"),
        ("时序规则命中策略", "all_match"),
    ):
        global_sheet.append([key, value])
    sheet = wb.create_sheet("时序提取规则")
    sheet.append([
        "是否启用", "规则名称", "工作簿关键字", "工作表关键字", "行头列", "列表头行",
        "数据起始行", "数据结束行", "数据起始列", "数据结束列",
    ])
    sheet.append(["是", "CompareRule", "sample", "Report", "A", "2,3", 5, "", "B", "C"])
    wb.save(path)


def python_baseline(template: Path, sources: list[Path], config: Path, output: Path) -> Path:
    sys.path.insert(0, str(PYTOOLS_SOURCE / "PyCode"))
    from core.compare_393 import run_compare
    from core.config_xlsx import GlobalConfig, load_timeline_rules
    from core.io_excel import clear_sheet_cache

    output.mkdir(parents=True, exist_ok=True)
    global_config = GlobalConfig(
        input_dir=template.parent,
        output_dir=output,
        log_dir=output / "logs",
        error_policy="continue",
        source_exts=[".xlsx"],
        timeline_rule_match_mode="all_match",
    )
    try:
        return run_compare(template, sources, load_timeline_rules(config), global_config)
    finally:
        clear_sheet_cache()
        logging.shutdown()


def run_rust(template: Path, sources: list[Path], config: Path) -> Path:
    command = (
        f"cd {shlex.quote(to_wsl_path(WORKSPACE))} && cargo run --quiet -p pytools-core --bin pytools -- "
        f"timeline-compare --config {shlex.quote(to_wsl_path(config))} "
        f"--template {shlex.quote(to_wsl_path(template))} "
        + " ".join(shlex.quote(to_wsl_path(path)) for path in sources)
    )
    completed = subprocess.run(["wsl", "-e", "sh", "-lc", command], check=True, capture_output=True)
    stdout = completed.stdout.decode("utf-8-sig", errors="replace").replace("\x00", "")
    stderr = completed.stderr.decode("utf-8-sig", errors="replace").replace("\x00", "")
    for line in reversed(stdout.splitlines()):
        try:
            payload = json.loads(line)
            return from_wsl_path(payload["output"])
        except (json.JSONDecodeError, KeyError):
            continue
    raise RuntimeError(f"Rust CLI did not return an output path: {stdout}\n{stderr}")


def read_values(path: Path) -> tuple[list[str], list[list[list[object]]]]:
    wb = load_workbook(path, data_only=False, read_only=True)
    names = []
    tables = []
    for ws in wb.worksheets:
        names.append(ws.title)
        tables.append([list(row) for row in ws.iter_rows(values_only=True)])
    wb.close()
    return names, tables


def main() -> int:
    if not PYTOOLS_SOURCE.is_dir():
        raise FileNotFoundError(PYTOOLS_SOURCE)
    with tempfile.TemporaryDirectory(prefix="pytools-timeline-compare-", ignore_cleanup_errors=True) as temp:
        root = Path(temp)
        inputs = root / "inputs"
        inputs.mkdir()
        template = root / "template.xlsx"
        unchanged = inputs / "sample_unchanged.xlsx"
        changed = inputs / "sample_changed.xlsx"
        unmatched = inputs / "unmatched.xlsx"
        make_book(template)
        make_book(unchanged)
        make_book(changed, changed=True)
        make_book(unmatched, changed=True)
        config = root / "config.xlsx"
        python_output = root / "python_output"
        rust_output = root / "rust_output"
        make_config(config, rust_output, inputs)
        sources = [unchanged, changed, unmatched, template]
        expected = python_baseline(template, sources, config, python_output)
        actual = run_rust(template, sources, config)
        expected_names, expected_values = read_values(expected)
        actual_names, actual_values = read_values(actual)
        if expected_names != actual_names or expected_values != actual_values:
            raise AssertionError(
                json.dumps(
                    {"expected": [expected_names, expected_values], "actual": [actual_names, actual_values]},
                    ensure_ascii=False,
                    default=str,
                    indent=2,
                )
            )
        print(json.dumps({"status": "passed", "sheets": expected_names, "sources": 3}, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
