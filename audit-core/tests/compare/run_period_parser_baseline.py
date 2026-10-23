"""Export period-parser vectors from the read-only Open-Data-Audit source."""

from __future__ import annotations

import argparse
import json
import platform
import sys
from pathlib import Path

sys.dont_write_bytecode = True

DEFAULT_SOURCE_ROOT = Path(r"C:\Users\AZI\Desktop\VibeCode\3.open-data-audit")
DEFAULT_OUTPUT = Path(__file__).resolve().parents[2] / "probe" / "reports" / "period_parser_python_baseline.json"

INPUTS = [
    "2026-06",
    "2026.6",
    "2026_06",
    "2026年6月",
    "2026/06",
    "2026-06-30",
    "2026-6-1",
    "2026.06.30",
    "2026_6_1",
    "2026/06/30",
    "2026年6月30日",
    "202606",
    "20260630",
    "20260101",
    "20261231",
    "202600",
    "202613",
    "20260230",
    "x2026-06y",
    "12026-06",
    "2026-060",
    "2026-6-32",
    "2026-06-32",
    "2026-06abc2026-07",
    "2026-06-01-2026-07-01",
    "2026-06-01x2026-07-01",
    "2026-00",
    "1999-06",
    "2099-12",
    "2026年13月",
    "2026年0月",
    "2026-06-1日",
    "2026-06-31日",
    "2026-06-3日",
    "2026-06月",
    "2026年06月报表",
    "报告2026.06.15汇总",
    "202606_202607",
    "2026-06-30-1",
    "",
    "١2026-06",
    "2026-06١",
    "20٢٦-٠٦",
    "20٢٦-06",
    "20٢٦06",
    "2026-٠٦",
    "2026061٢",
    "²2026-06",
    "一2026-06",
]


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--source-root", type=Path, default=DEFAULT_SOURCE_ROOT)
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    args = parser.parse_args()

    source_package = args.source_root / "core" / "src"
    if not source_package.is_dir():
        parser.error(f"Open-Data-Audit source not found: {source_package}")
    sys.path.insert(0, str(source_package))
    from base_audit.discovery import _periods_in_text

    report = {
        "schema_version": 1,
        "source_project": "Open-Data-Audit",
        "source_function": "base_audit.discovery._periods_in_text",
        "source_file": "core/src/base_audit/discovery.py",
        "python_version": platform.python_version(),
        "case_count": len(INPUTS),
        "cases": [
            {"input": value, "periods": sorted(_periods_in_text(value))}
            for value in INPUTS
        ],
        "scope": "Pure text parser only; no directory scan or workbook I/O.",
    }
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(f"wrote {len(INPUTS)} source-baseline cases to {args.output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
