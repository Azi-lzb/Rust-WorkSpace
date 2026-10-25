#!/usr/bin/env python3
"""Compare the Rust S5 run-outputs port with the Python write_run_outputs.

Both sides run the full enabled 单位贷款 rule set against the same inputs and
write outputs under parallel temporary roots with a fixed run id (Python uses
`engine.run(output_dir=...)` with the openpyxl writer engine as the styling
baseline). The comparator diffs: audit copies (byte-level), the manifest
(path-normalized), the 审核结果.xlsx workbooks (values/types, hyperlinks,
widths, freeze, filter, row heights, header/body styles, numeric formats,
page setup, sheet inventory), and the ledger's stable fields. Aggregate
counts only.
"""

from __future__ import annotations

import argparse
import json
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

from openpyxl.utils import column_index_from_string, get_column_letter

ROOT = Path(__file__).resolve().parents[3]
RUN_ID = "PARITYRUN"


def normalize_link(target: str, root: Path) -> str:
    text = str(target).replace("\\", "/")
    # Resolve relative copy links to (basename, fragment) pairs.
    if "#" in text:
        path_part, fragment = text.split("#", 1)
        basename = Path(path_part).name
        return f"{basename}#{fragment}"
    return Path(text).name


def normalize_value(value):
    if value is None or value == "":
        return None
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        number = float(value)
        if number.is_integer() and abs(number) < 1e16:
            return f"{number:.1f}"
        return repr(number)
    if isinstance(value, str):
        try:
            number = float(value)
        except ValueError:
            return value
        if value.strip() != value:
            return value
        return value
    return value


def expand_widths(dimensions):
    """Expand grouped <col min max> spans (openpyxl records them under the
    first letter only) into per-letter widths."""
    widths = {}
    for dimension in dimensions.values():
        if dimension.width is None:
            continue
        min_col = dimension.min or column_index_from_string(dimension.index)
        max_col = dimension.max or min_col
        for index in range(min_col, max_col + 1):
            widths[get_column_letter(index)] = round(dimension.width, 6)
    return widths

def normalize_color(rgb):
    """Colors compare modulo the 8-digit alpha prefix: openpyxl writes 00/FF
    prefixes depending on the producing library while Excel semantics only
    cover the six-digit RGB."""
    if not rgb:
        return rgb
    text = str(rgb)
    return text[-6:]


def normalize_fit(value, fit_to_page):
    """fitToWidth=1 is the effective default under fitToPage; rust_xlsxwriter
    omits the attribute when it equals 1 while the openpyxl baseline writes it
    explicitly."""
    if fit_to_page and value is None:
        return 1
    return value

def read_workbook_facts(path: Path):
    from openpyxl import load_workbook

    workbook = load_workbook(path, data_only=False)
    facts = {}
    for sheet in workbook.worksheets:
        sheet_facts = {
            "freeze": sheet.freeze_panes,
            "filter": sheet.auto_filter.ref,
            "gridlines": sheet.sheet_view.showGridLines,
            "orientation": sheet.page_setup.orientation,
            "paper_size": sheet.page_setup.paperSize,
            "fit_to_width": normalize_fit(
                sheet.page_setup.fitToWidth,
                bool(
                    sheet.sheet_properties.pageSetUpPr
                    and sheet.sheet_properties.pageSetUpPr.fitToPage
                ),
            ),
            "fit_to_height": sheet.page_setup.fitToHeight,
            "print_titles": sheet.print_title_rows,
            "widths": expand_widths(sheet.column_dimensions),
            "row_heights": {
                key: value.height
                for key, value in sheet.row_dimensions.items()
                if value.height is not None
            },
            "header_cells": [],
            "cells": [],
            "hyperlinks": {},
        }
        header = next(sheet.iter_rows(min_row=1, max_row=1), ())
        for cell in header:
            font = cell.font
            fill = cell.fill
            sheet_facts["header_cells"].append({
                "value": cell.value,
                "font": (
                    font.name if font else None,
                    font.size if font else None,
                    bool(font.bold) if font else None,
                    normalize_color(font.color.rgb if font and font.color else None),
                ),
                "fill": (
                    normalize_color(fill.fgColor.rgb if fill and fill.fgColor else None),
                ),
                "align": (
                    cell.alignment.horizontal,
                    cell.alignment.vertical,
                    bool(cell.alignment.wrap_text),
                ),
            })
        for row in sheet.iter_rows(min_row=2):
            for cell in row:
                if cell.value is None and cell.hyperlink is None:
                    continue
                entry = {
                    "ref": cell.coordinate,
                    "value": normalize_value(cell.value),
                    "type": cell.data_type,
                    "format": cell.number_format,
                }
                if cell.hyperlink is not None:
                    entry["link"] = normalize_link(cell.hyperlink.target, path.parent)
                if cell.value is not None:
                    font = cell.font
                    entry["font"] = (
                        font.name if font else None,
                        font.size if font else None,
                    )
                    entry["align"] = (
                        cell.alignment.horizontal,
                        cell.alignment.vertical,
                        bool(cell.alignment.wrap_text),
                    )
                sheet_facts["cells"].append(entry)
        facts[sheet.title] = sheet_facts
    workbook.close()
    return facts


def diff(left, right, path: str, diffs: list[str]) -> None:
    if isinstance(left, dict) and isinstance(right, dict):
        for key in sorted(set(left) | set(right)):
            diff(left.get(key), right.get(key), f"{path}.{key}", diffs)
    elif isinstance(left, list) and isinstance(right, list):
        if len(left) != len(right):
            diffs.append(f"{path}: length {len(left)} vs {len(right)}")
        for index in range(min(len(left), len(right))):
            diff(left[index], right[index], f"{path}[{index}]", diffs)
    elif left != right:
        diffs.append(path)


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--source-root", type=Path, required=True)
    parser.add_argument("--config-workbook", type=Path, required=True)
    parser.add_argument("--audit-copy", type=Path, required=True)
    parser.add_argument("--template-path", type=Path)
    parser.add_argument("--template-dir", type=Path)
    parser.add_argument("--workspace", type=Path, default=ROOT)
    parser.add_argument("--report", type=Path, required=True)
    args = parser.parse_args()

    source_root = args.source_root.resolve(strict=True)
    workspace = args.workspace.resolve(strict=True)
    report_path = args.report.resolve()
    template_path = args.template_path.resolve(strict=True) if args.template_path else None
    template_dir = args.template_dir.resolve(strict=True) if args.template_dir else None
    sys.path.insert(0, str(source_root / "core" / "src"))
    os.environ["AUDIT3_WRITER_ENGINE"] = "openpyxl"

    with tempfile.TemporaryDirectory(prefix="audit-s5-outputs-") as temporary:
        directory = Path(temporary)
        config_copy = directory / "config.xlsx"
        source_copy = directory / "source.xlsx"
        shutil.copy2(args.config_workbook.resolve(strict=True), config_copy)
        shutil.copy2(args.audit_copy.resolve(strict=True), source_copy)
        py_root = directory / "py_out"
        rs_root = directory / "rs_out"

        # --- Python side: full run + write_run_outputs ---
        from s1audit_configured.engine import AuditEngine
        from s1audit_configured.model import RuntimeContext, SourceBinding
        from s1audit_configured.runtime import RuntimePaths

        engine = AuditEngine(config_copy)
        context = RuntimeContext(
            period="2026-07-31",
            institution_code="PARITY",
            institution_name="PARITY",
            current_sources={
                "CURRENT": SourceBinding(
                    "CURRENT",
                    source_copy,
                    business="单位贷款",
                    period="2026-07-31",
                    institution_code="PARITY",
                )
            },
        )
        result = engine.run(
            context,
            output_dir=py_root,
            run_id=RUN_ID,
            businesses=["单位贷款"],
            schema_snapshot_dir=directory / "snap",
            persist_schema_snapshots=False,
            runtime_paths=RuntimePaths(
                template_path=template_path,
                template_dir=template_dir,
            ),
        )
        py_run_dir = py_root / RUN_ID

        # --- Rust side ---
        request = {
            "config": str(config_copy),
            "source": str(source_copy),
            "template_path": str(template_path) if template_path else None,
            "template_dir": str(template_dir) if template_dir else None,
            "period": "2026-07-31",
            "institution_code": "PARITY",
            "institution_name": "PARITY",
            "source_id": "CURRENT",
            "run_id": RUN_ID,
            "output_root": str(rs_root),
        }
        input_path = directory / "request.json"
        input_path.write_text(json.dumps(request, ensure_ascii=False), encoding="utf-8")
        rust_report = directory / "rust_summary.json"
        environment = os.environ.copy()
        environment["AUDIT_S5_OUTPUTS_INPUT"] = str(input_path)
        environment["AUDIT_S5_OUTPUTS_REPORT"] = str(rust_report)
        completed = subprocess.run(
            [
                "cargo", "test", "-p", "audit-core", "--test", "s5_outputs_parity",
                "--locked", "--", "--ignored", "--nocapture",
            ],
            cwd=workspace,
            env=environment,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            check=False,
        )
        if completed.returncode != 0 or not rust_report.is_file():
            print(json.dumps({
                "status": "rust_runner_failed",
                "exit_code": completed.returncode,
                "stderr_tail": completed.stderr[-1500:],
            }))
            return 1
        summary = json.loads(rust_report.read_text(encoding="utf-8"))
        rs_run_dir = Path(summary["run_dir"])

        # --- compare artifacts ---
        diffs: list[str] = []

        # 1. audit copies: same file set, byte-identical
        py_copies = sorted(p.name for p in (py_run_dir / "审核副本").iterdir())
        rs_copies = sorted(p.name for p in (rs_run_dir / "审核副本").iterdir())
        if py_copies != rs_copies:
            diffs.append(f"copies.set: {py_copies} vs {rs_copies}")
        else:
            for name in py_copies:
                if (py_run_dir / "审核副本" / name).read_bytes() != (
                    rs_run_dir / "审核副本" / name
                ).read_bytes():
                    diffs.append(f"copies.bytes:{name}")

        # 2. workbooks
        py_book = read_workbook_facts(py_run_dir / "审核结果.xlsx")
        rs_book = read_workbook_facts(rs_run_dir / "审核结果.xlsx")
        diff(py_book, rs_book, "workbook", diffs)

        # 3. ledger stable fields
        py_ledger = json.loads((py_run_dir / "运行总账.json").read_text(encoding="utf-8"))
        rs_ledger = json.loads((rs_run_dir / "运行总账.json").read_text(encoding="utf-8"))
        stable = {
            "run_id": (py_ledger.get("run_id"), rs_ledger.get("run_id")),
            "status_counts": (py_ledger.get("status_counts"), rs_ledger.get("status_counts")),
            "run_status": (py_ledger.get("run_status"), rs_ledger.get("run_status")),
        }
        for field, (left, right) in stable.items():
            if left != right:
                diffs.append(f"ledger.{field}")

        # 4. manifest: roles/order/snapshot flags (paths normalized to names)
        def manifest_facts(ledger):
            facts = []
            for entry in ledger.get("audit_copy_manifest", []):
                facts.append({
                    "role": entry.get("role"),
                    "source_id": entry.get("source_id"),
                    "copy_name": Path(entry.get("copy_path", "")).name,
                    "source_name": Path(entry.get("source_path", "")).name,
                    "matches_snapshot": entry.get("matches_snapshot"),
                    "sha_before": entry.get("sha256_before"),
                    "sha_after": entry.get("sha256_after"),
                    "sha_copy": entry.get("sha256_copy"),
                })
            return facts

        diff(
            manifest_facts(py_ledger),
            manifest_facts(rs_ledger),
            "ledger.manifest",
            diffs,
        )

        triggered = summary.get("triggered_rows")
        failures = summary.get("failure_rows")
        print(json.dumps({
            "rust_triggered_rows": triggered,
            "rust_failure_rows": failures,
            "py_status_counts": py_ledger.get("status_counts"),
            "diff_count": len(diffs),
            "diff_sample": sorted(set(diffs))[:15],
        }, ensure_ascii=False))

    report = {
        "schema_version": 1,
        "suite": "audit.s5_run_outputs_parity",
        "status": "completed" if not diffs else "mismatch",
        "sample_scope": {
            "business": "单位贷款",
            "run_id": RUN_ID,
            "files_copied_to_temporary": True,
            "template_path_supplied": template_path is not None,
            "template_dir_supplied": template_dir is not None,
            "python_writer_engine": "openpyxl (styling baseline)",
        },
        "comparison": {
            "diff_count": len(diffs),
            "diff_paths_sample": sorted(set(diffs))[:20],
        },
        "verification": {
            "rust_runner": "audit-core/tests/s5_outputs_parity.rs",
            "python_baseline": (
                "s1audit_configured.engine.run(output_dir=...) with "
                "AUDIT3_WRITER_ENGINE=openpyxl"
            ),
            "python_baseline_version": sys.version.split()[0],
            "inputs_unchanged": True,
            "raw_values_persisted": False,
            "capability_status": "unverified",
            "production_route_eligible": False,
        },
        "limitations": [
            "Conditional-format and data-quality tables are compared empty; "
            "their populated paths belong to the S1 conditional-format bridge.",
            "Structure snapshots and hash-integrity downgrade are out of scope.",
            "Colors compare modulo the 8-digit alpha prefix and fitToWidth=1 "
            "compares equal to an omitted attribute under fitToPage: both are "
            "rust_xlsxwriter/openpyxl serializer representation deltas with "
            "identical Excel semantics.",
        ],
    }
    report_path.parent.mkdir(parents=True, exist_ok=True)
    report_path.write_text(
        json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )
    print(json.dumps({"status": report["status"], "diffs": len(diffs)}))
    return 0 if not diffs else 1


if __name__ == "__main__":
    raise SystemExit(main())
