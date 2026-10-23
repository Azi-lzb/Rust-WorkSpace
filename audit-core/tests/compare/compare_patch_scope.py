"""Confirm the experimental OOXML repair only changes worksheet column definitions."""
from __future__ import annotations

import argparse
import json
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"


def without_cols(data: bytes) -> bytes:
    root = ET.fromstring(data)
    for child in list(root):
        if child.tag == f"{{{MAIN_NS}}}cols":
            root.remove(child)
    return ET.tostring(root, encoding="utf-8")


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--raw", type=Path, required=True)
    parser.add_argument("--patched", type=Path, required=True)
    parser.add_argument("--report", type=Path)
    args = parser.parse_args()
    checks = []
    with ZipFile(args.raw) as raw, ZipFile(args.patched) as patched:
        raw_names, patched_names = set(raw.namelist()), set(patched.namelist())
        checks.append({"name": "zip_parts_unchanged", "passed": raw_names == patched_names, "detail": f"raw={len(raw_names)}; patched={len(patched_names)}"})
        changed_parts = []
        failed_parts = []
        for name in sorted(raw_names & patched_names):
            before, after = raw.read(name), patched.read(name)
            if before == after:
                continue
            changed_parts.append(name)
            if name.startswith("xl/worksheets/") and name.endswith(".xml"):
                if without_cols(before) != without_cols(after):
                    failed_parts.append(name)
            else:
                failed_parts.append(name)
        checks.append({"name": "change_scope_is_worksheet_cols", "passed": not failed_parts and bool(changed_parts), "changed_parts": changed_parts, "unexpected_parts": failed_parts})
        worksheet_changes = []
        for name in changed_parts:
            if name.startswith("xl/worksheets/") and name.endswith(".xml"):
                b_root, a_root = ET.fromstring(raw.read(name)), ET.fromstring(patched.read(name))
                b_cols = b_root.find(f"{{{MAIN_NS}}}cols")
                a_cols = a_root.find(f"{{{MAIN_NS}}}cols")
                worksheet_changes.append({"part": name, "raw_cols": None if b_cols is None else ET.tostring(b_cols, encoding="unicode"), "patched_cols": None if a_cols is None else ET.tostring(a_cols, encoding="unicode")})
        checks.append({"name": "column_definitions_logged", "passed": bool(worksheet_changes), "details": worksheet_changes})
    passed = sum(c["passed"] for c in checks)
    result = {"raw": str(args.raw), "patched": str(args.patched), "checks": checks, "passed": passed, "total": len(checks), "rate": passed / len(checks)}
    serialized = json.dumps(result, ensure_ascii=False, indent=2)
    if args.report:
        args.report.parent.mkdir(parents=True, exist_ok=True)
        args.report.write_text(serialized + "\n", encoding="utf-8")
    print(serialized)


if __name__ == "__main__":
    main()
