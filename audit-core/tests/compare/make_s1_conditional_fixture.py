#!/usr/bin/env python3
"""Generate a synthetic XLSX for the Rust S1 conditional-evaluator probe."""

from __future__ import annotations

import argparse
from pathlib import Path

from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule, ColorScaleRule, FormulaRule


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--output", required=True, type=Path)
    output = parser.parse_args().output
    output.parent.mkdir(parents=True, exist_ok=True)

    book = Workbook()
    sheet = book.active
    sheet.title = "S1Conditional"
    values = {
        "B2": 5,
        "B3": -1,
        "B4": 0,
        "C5": 2,
        "C6": 0,
        "C7": 3,
        "A1": 1,
        "A2": 2,
        "A3": 3,
        "C2": 6,
        "D27": 31894.55,
        "D28": 31.28,
        "D34": 0,
        "D35": 31925.83,
    }
    for address, value in values.items():
        sheet[address] = value

    sheet.conditional_formatting.add(
        "B2:B4", CellIsRule(operator="greaterThan", formula=["0"])
    )
    sheet.conditional_formatting.add(
        "C5:C7", FormulaRule(formula=["AND(C5>0,C5>1)"])
    )
    sheet.conditional_formatting.add(
        "E2", FormulaRule(formula=["$A$1+$A$2+$A$3=$C$2"])
    )
    sheet.conditional_formatting.add(
        "E35", FormulaRule(formula=["$D35<>$D27+$D28+$D34"])
    )
    sheet.conditional_formatting.add(
        "F5", FormulaRule(formula=['INDIRECT("A1")>0'])
    )
    sheet.conditional_formatting.add("G5", FormulaRule(formula=["((("]))
    sheet.conditional_formatting.add(
        "H2:H3",
        ColorScaleRule(
            start_type="num",
            start_value=0,
            start_color="FFFF0000",
            end_type="num",
            end_value=10,
            end_color="FF00FF00",
        ),
    )
    book.save(output)
    book.close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
