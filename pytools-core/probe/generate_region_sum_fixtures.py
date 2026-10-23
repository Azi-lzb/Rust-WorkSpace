"""Create deterministic synthetic workbooks for the region-sum feature."""

from __future__ import annotations

import json
from pathlib import Path

from openpyxl import Workbook


FIXTURE_DIR = Path(__file__).parents[1] / "tests" / "fixtures" / "region_sum"


def save_book(path: Path, sheets: dict[str, list[list[object]]]) -> None:
    book = Workbook()
    book.remove(book.active)
    for title, rows in sheets.items():
        sheet = book.create_sheet(title)
        for row in rows:
            sheet.append(row)
    path.parent.mkdir(parents=True, exist_ok=True)
    book.save(path)
    book.close()


def main() -> None:
    save_book(
        FIXTURE_DIR / "balanced.xlsx",
        {
            "本期分地区": [
                ["惠州市分地区数据"],
                ["序号", "地区", "指标甲", "指标乙", "指标丙"],
                [None, "惠州市", 100, 1000, 0],
                [1, "惠阳区", 10, 100, 0],
                [2, "博罗县", 20, 200, 0],
                [3, "惠东县", 15, 150, 0],
                [4, "龙门县", 10, 100, 0],
                [5, "惠州市本部", 20, 200, 0],
                [6, "大亚湾", 25, 250, 0],
                [None, "惠州市", 999, 999, 999],
            ]
        },
    )
    save_book(
        FIXTURE_DIR / "mismatch.xlsx",
        {
            "分地区汇总": [
                ["地区总分校验"],
                ["序号", "地区", "金额", "比例", "非数值"],
                [None, "惠州市", 100, 1000, "bad"],
                [1, "惠阳区", 40, 500, "bad"],
                [2, "博罗县", 50, 500, "bad"],
            ]
        },
    )
    save_book(FIXTURE_DIR / "no_target.xlsx", {"普通表": [["地区", "金额"], ["惠州市", 100]]})
    save_book(
        FIXTURE_DIR / "non_numeric_missing.xlsx",
        {
            "年度分地区": [
                ["年度数据"],
                ["序号", "地区", "金额"],
                [None, "惠州市", "bad"],
                [1, "惠阳区", "bad"],
            ]
        },
    )
    (FIXTURE_DIR / "broken.xlsx").write_bytes(b"not a workbook")
    cases = {
        "schema_version": 1,
        "run_files": [
            "balanced.xlsx",
            "mismatch.xlsx",
            "no_target.xlsx",
            "non_numeric_missing.xlsx",
            "broken.xlsx",
        ],
        "expected_stats": {"files": 4, "sheets": 3, "errors": 1, "failed_files": 1},
    }
    (FIXTURE_DIR / "cases.json").write_text(
        json.dumps(cases, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )
    print(json.dumps({"fixture_dir": str(FIXTURE_DIR), "case_count": len(cases["run_files"])}, ensure_ascii=False))


if __name__ == "__main__":
    main()
