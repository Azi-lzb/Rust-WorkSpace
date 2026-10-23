"""Run the original Python template recommender on synthetic workbook cases."""

from __future__ import annotations

import argparse
import hashlib
import json
import sys
import tempfile
import zipfile
from pathlib import Path


WORKSPACE = Path(__file__).resolve().parents[2]
CASES = WORKSPACE / "tests/fixtures/template_recommendation_cases.json"
WORKBOOK_XML = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheets>{sheets}</sheets>
</workbook>'''


def write_sheet_names(path: Path, names: list[str]) -> None:
    sheets = "".join(
        f'<sheet name="{name}" sheetId="{index}"/>'
        for index, name in enumerate(names, start=1)
    )
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr("xl/workbook.xml", WORKBOOK_XML.format(sheets=sheets))


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--source-root",
        type=Path,
        default=Path(r"C:\Users\AZI\Desktop\VibeCode\3.open-data-audit"),
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=WORKSPACE / "probe/reports/template_recommendation_python_baseline.json",
    )
    args = parser.parse_args()
    sys.dont_write_bytecode = True
    sys.path.insert(0, str(args.source_root / "core/src"))

    from base_audit.discovery import recommend_template

    cases = json.loads(CASES.read_text(encoding="utf-8"))["cases"]
    source_file = args.source_root / "core/src/base_audit/discovery.py"
    results = []
    with tempfile.TemporaryDirectory(prefix="audit-template-baseline-") as temp:
        root = Path(temp)
        for case in cases:
            template_dir = root / case["id"] / "templates"
            input_dir = root / case["id"] / "input"
            template_dir.mkdir(parents=True)
            input_dir.mkdir(parents=True)
            for item in case["templates"]:
                write_sheet_names(template_dir / item["name"], item["sheets"])
            for index, names in enumerate(case["sources"], start=1):
                write_sheet_names(input_dir / f"source-{index}.xlsx", names)

            suggestion = recommend_template(
                template_dir,
                input_dir,
                root / case["id"] / "cache" / "template-index.json",
            )
            results.append(
                {
                    "id": case["id"],
                    "result": {
                        "template": (
                            suggestion.template_path.name
                            if suggestion.template_path is not None
                            else None
                        ),
                        "confidence": suggestion.confidence,
                        "matched": suggestion.matched,
                        "details": suggestion.details,
                        "alternatives": [list(item) for item in suggestion.alternatives],
                    },
                }
            )

    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(
        json.dumps(
            {
                "schema_version": 1,
                "suite": "audit.s1_template_recommendation",
                "source_project": "Open-Data-Audit",
                "source": {
                    "module": "core/src/base_audit/discovery.py",
                    "sha256": hashlib.sha256(source_file.read_bytes()).hexdigest().upper(),
                    "python_version": sys.version.split()[0],
                },
                "source_function": "base_audit.discovery.recommend_template",
                "cases": results,
            },
            ensure_ascii=False,
            indent=2,
        )
        + "\n",
        encoding="utf-8",
    )
    print(args.output)


if __name__ == "__main__":
    main()
