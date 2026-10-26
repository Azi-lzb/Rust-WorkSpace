"""Compare the real PyCode-configured wide rules through the interactive Rust menu.

The configuration sheets are copied read-only from the public VBA toolbox fixture;
source workbooks are synthetic and generated under a temporary directory. This
closes the menu-dispatch gap only and is not real PyTools report acceptance.
"""
from __future__ import annotations

import datetime as dt
import json
import os
import subprocess
import sys
import tempfile
from pathlib import Path

sys.dont_write_bytecode = True
PROBE = Path(__file__).resolve().parent
WORKSPACE = PROBE.parents[1]
sys.path.insert(0, str(PROBE))

import compare_real_rules_wide as real_rules  # noqa: E402
import compare_wide_summary as wide_probe  # noqa: E402


def run_menu(
    rust_cli: Path,
    config: Path,
    sources: list[Path],
    menu_choice: str,
) -> subprocess.CompletedProcess[str]:
    source_line = ";".join(str(path) for path in sources)
    input_text = f"1\n{menu_choice}\n{source_line}\n0\n"
    completed = subprocess.run(
        [str(rust_cli), "menu", "--config", str(config), "--log-view", "developer"],
        cwd=WORKSPACE,
        input=input_text,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        timeout=120,
        check=False,
    )
    if completed.returncode != 0:
        raise AssertionError(
            f"menu route 1->{menu_choice} exited {completed.returncode}: "
            f"{completed.stderr}\n{completed.stdout}"
        )
    if "宽表汇总" not in completed.stdout:
        raise AssertionError(f"menu route 1->{menu_choice} did not run wide summary")
    return completed


def run_case(
    root: Path,
    rust_cli: Path,
    sources: list[Path],
    row_suffix: bool,
) -> dict[str, object]:
    feature_id = "PT-1-3" if row_suffix else "PT-1-4"
    menu_choice = "3" if row_suffix else "4"
    case = root / feature_id
    python_output_dir = case / "python-output"
    rust_output_dir = case / "rust-output"
    python_config = case / "python-config.xlsx"
    rust_config = case / "rust-config.xlsx"
    real_rules.make_bridge_config(python_config, python_output_dir)
    real_rules.make_bridge_config(rust_config, rust_output_dir)

    python_output = real_rules.python_side(python_config, sources, row_suffix)
    if python_output is None or not python_output.is_file():
        raise AssertionError(f"PyCode produced no {feature_id} workbook")

    menu_result = run_menu(rust_cli, rust_config, sources, menu_choice)
    rust_outputs = sorted(
        path for path in rust_output_dir.glob("*.xlsx") if path.is_file()
    )
    if len(rust_outputs) != 1:
        raise AssertionError(
            f"expected one Rust {feature_id} output from menu, found {len(rust_outputs)}"
        )
    rust_output = rust_outputs[0]
    wide_probe.assert_parity(python_output, rust_output)

    log_dir = rust_output_dir / "logs"
    histories = sorted(log_dir.glob(f"{feature_id}_*.json"))
    if len(histories) != 1:
        raise AssertionError(
            f"expected one developer history for {feature_id}, found {len(histories)}"
        )
    history = json.loads(histories[0].read_text(encoding="utf-8"))
    engines = history.get("actual_engines", [])
    if history.get("status") != "succeeded":
        raise AssertionError(f"{feature_id} menu history did not succeed")
    if history.get("completed_inputs") != len(sources) or history.get("failed_inputs") != 0:
        raise AssertionError(f"{feature_id} menu history input counts differ: {history}")
    if not {"calamine", "rust_xlsxwriter"}.issubset(set(engines)):
        raise AssertionError(f"{feature_id} menu history engine record is incomplete: {engines}")
    history_text = json.dumps(history, ensure_ascii=False)
    if any(str(source) in history_text for source in sources):
        raise AssertionError(f"{feature_id} developer history contains source paths")

    projection = wide_probe.normalized(rust_output)
    return {
        "feature_id": feature_id,
        "menu_route": f"pytools menu -> 1 -> {menu_choice}",
        "status": "passed_exact",
        "processed_sources": len(sources),
        "python_rust_projection": "exact",
        "compared_fields": [
            "sheet_names_and_order",
            "dimensions",
            "cell_values_and_types",
            "number_formats",
            "effective_styles",
            "merged_ranges",
            "formulas",
            "defined_names",
            "freeze_panes",
        ],
        "output_sheet_names": projection["sheet_names"],
        "output_dimensions": projection["dimensions"],
        "actual_engines": engines,
        "developer_history": "succeeded; source paths omitted",
        "target_append": "not exercised; bridge configuration has no target paths",
    }


def main() -> int:
    rust_cli_value = os.environ.get("PYTOOLS_RUST_CLI")
    temp_root_value = os.environ.get("PYTOOLS_PROBE_TEMP")
    if not rust_cli_value:
        raise RuntimeError("set PYTOOLS_RUST_CLI to the native Windows pytools.exe")
    if not temp_root_value:
        raise RuntimeError("set PYTOOLS_PROBE_TEMP to a temporary directory on F:")
    rust_cli = Path(rust_cli_value)
    temp_root = Path(temp_root_value)
    if not rust_cli.is_file():
        raise FileNotFoundError(rust_cli)
    if not temp_root.is_dir():
        raise NotADirectoryError(temp_root)

    with tempfile.TemporaryDirectory(
        prefix="pytools-real-rules-menu-", dir=temp_root, ignore_cleanup_errors=True
    ) as temp:
        root = Path(temp)
        sources = [
            root / "sources" / "20260609_月报一批分地区.xlsx",
            root / "sources" / "20260609_绿色普惠分地区.xlsx",
            root / "sources" / "2026分地区统计研究调研问卷.xlsx",
        ]
        real_rules.make_monthly_source(sources[0])
        real_rules.make_inclusive_source(sources[1])
        real_rules.make_survey_source(sources[2])
        source_before = {path: path.read_bytes() for path in sources}
        cases = [run_case(root, rust_cli, sources, row_suffix) for row_suffix in (True, False)]
        if any(path.read_bytes() != source_before[path] for path in sources):
            raise AssertionError("menu processing modified a synthetic source workbook")

    report = {
        "schema_version": 1,
        "suite": "pytools.real_rules_interactive_menu_wide_summary_parity",
        "date": dt.date.today().isoformat(),
        "status": "passed_synthetic_sources_real_rules_menu_parity",
        "environment": {
            "platform": "Windows 11 x64",
            "rust_cli": "native MSVC executable",
            "sample_kind": "synthetic source workbooks",
        },
        "configuration_source": {
            "workbook": "tests/fixtures/vba_config/VbaV2.4.2.xlsm",
            "origin": "public VBA_Excel_Tools fixture, HEAD cea6dfa",
            "scope": "15 real timeline rules and 146 real path maps copied into temporary bridge configs; global paths synthetic",
        },
        "entry": "pytools menu -> 1 -> 3/4",
        "source_workbooks": "three synthetic workbooks shaped to hit three real rules",
        "cases": cases,
        "source_workbooks_unchanged": True,
        "limitations": [
            "Source workbooks are synthetic, not original PyTools business reports; this is not real-sample acceptance.",
            "The menu used manual semicolon-separated paths; native file-dialog UI was not exercised.",
            "No configured target append was exercised; no production ExcelModifier capability was used.",
            "Kylin V10, UOS V20, glibc 2.28, and offline deb delivery remain unverified.",
        ],
    }
    report_path = PROBE / "reports" / "real_rules_menu_parity_20261026.json"
    report_path.parent.mkdir(parents=True, exist_ok=True)
    report_path.write_text(
        json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )
    print(
        "REAL_RULES_MENU_RESULT="
        + json.dumps(
            {
                "status": report["status"],
                "cases": [
                    {
                        "feature_id": case["feature_id"],
                        "status": case["status"],
                        "dimensions": case["output_dimensions"],
                        "engines": case["actual_engines"],
                    }
                    for case in cases
                ],
                "report": "pytools-core/probe/reports/real_rules_menu_parity_20261026.json",
            },
            ensure_ascii=False,
        )
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
