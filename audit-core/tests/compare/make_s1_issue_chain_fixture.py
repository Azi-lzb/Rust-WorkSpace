#!/usr/bin/env python3
"""Generate a minimal synthetic workbook for the Rust S1 issue-chain CLI test."""

from __future__ import annotations

import argparse
from pathlib import Path

from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--output", required=True, type=Path)
    output = parser.parse_args().output
    output.parent.mkdir(parents=True, exist_ok=True)

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "S1Conditional"
    sheet["A2"] = "资产总额"
    sheet["B2"] = 5
    sheet.conditional_formatting.add(
        "B2", CellIsRule(operator="greaterThan", formula=["0"])
    )
    workbook.save(output)
    workbook.close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
