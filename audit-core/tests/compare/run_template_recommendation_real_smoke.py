"""Read-only smoke comparison against real Audit templates and source files."""

from __future__ import annotations

import argparse
import hashlib
import json
import sys
import tempfile
from pathlib import Path


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--source-root",
        type=Path,
        default=Path(r"C:\Users\AZI\Desktop\VibeCode\3.open-data-audit"),
    )
    parser.add_argument("--template-dir", type=Path, required=True)
    parser.add_argument("--input-dir", type=Path, required=True)
    parser.add_argument("--recursive", action="store_true")
    args = parser.parse_args()
    sys.dont_write_bytecode = True
    sys.path.insert(0, str(args.source_root / "core/src"))

    from base_audit.discovery import recommend_template, source_workbooks

    source_file = args.source_root / "core/src/base_audit/discovery.py"

    with tempfile.TemporaryDirectory(prefix="audit-template-index-") as temp:
        suggestion = recommend_template(
            args.template_dir,
            args.input_dir,
            Path(temp) / "template-index.json",
            recursive=args.recursive,
        )
    payload = {
        "source_python_version": sys.version.split()[0],
        "source_discovery_sha256": hashlib.sha256(source_file.read_bytes()).hexdigest().upper(),
        "source_workbook_count": len(
            source_workbooks(args.input_dir, recursive=args.recursive)
        ),
        "template": (
            suggestion.template_path.name
            if suggestion.template_path is not None
            else None
        ),
        "confidence": suggestion.confidence,
        "matched": suggestion.matched,
        "details": suggestion.details,
        "alternatives": [list(item) for item in suggestion.alternatives],
    }
    print(json.dumps(payload, ensure_ascii=False))


if __name__ == "__main__":
    main()
