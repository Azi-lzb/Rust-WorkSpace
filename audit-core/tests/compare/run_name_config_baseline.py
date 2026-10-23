"""Generate a read-only baseline for pure base_audit.name_config helpers."""

from __future__ import annotations

import argparse
import hashlib
import json
import sys
from dataclasses import asdict
from pathlib import Path


WORKSPACE = Path(__file__).resolve().parents[2]
CASES = WORKSPACE / "tests/fixtures/name_config_helper_cases.json"
SOURCE_RELATIVE = Path("core/src/base_audit/name_config.py")


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
        default=WORKSPACE / "probe/reports/name_config_helpers_python_baseline.json",
    )
    args = parser.parse_args()
    sys.dont_write_bytecode = True
    sys.path.insert(0, str(args.source_root / "core/src"))

    from base_audit.name_config import (
        FeatureMapping,
        _limited,
        _split_names,
        features_of_type,
        matches_named_range,
    )

    cases = json.loads(CASES.read_text(encoding="utf-8"))
    mappings = [FeatureMapping(**item) for item in cases["feature_mappings"]]
    payload = {
        "source": {
            "project": str(args.source_root),
            "module": SOURCE_RELATIVE.as_posix(),
            "sha256": hashlib.sha256(
                (args.source_root / SOURCE_RELATIVE).read_bytes()
            ).hexdigest().upper(),
        },
        "feature_mappings": [asdict(mapping) for mapping in mappings],
        "split_names": [
            {"input": value, "output": list(_split_names(value))}
            for value in cases["split_names"]
        ],
        "limited_values": [
            {"input": value, "output": _limited(value)}
            for value in cases["limited_values"]
        ],
        "features_of_type": [],
        "named_range_cases": [
            {
                **case,
                "output": matches_named_range(
                    mappings[case["mapping_index"]], case["candidate"]
                ),
            }
            for case in cases["named_range_cases"]
        ],
    }
    for feature_type in cases["features_of_type"]:
        selected = features_of_type(mappings, feature_type)
        selected_ids = {id(feature) for feature in selected}
        payload["features_of_type"].append(
            {
                "feature_type": feature_type,
                "indices": [
                    index
                    for index, feature in enumerate(mappings)
                    if id(feature) in selected_ids
                ],
            }
        )
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )
    print(args.output)


if __name__ == "__main__":
    main()
