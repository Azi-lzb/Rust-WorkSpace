"""Measure the probe process's internal elapsed_ms after one warm-up run."""
from __future__ import annotations

import argparse
import json
import statistics
import subprocess
from pathlib import Path


def wsl_path(path: Path) -> str:
    text = str(path.resolve())
    if len(text) >= 2 and text[1] == ":":
        return f"/mnt/{text[0].lower()}/{text[3:].replace(chr(92), '/') }"
    return text.replace("\\", "/")


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--probe-bin", required=True)
    parser.add_argument("--wsl", action="store_true")
    parser.add_argument("--input", type=Path, required=True)
    parser.add_argument("--output-dir", type=Path, required=True)
    parser.add_argument("--report", type=Path, required=True)
    parser.add_argument("--runs", type=int, default=5)
    args = parser.parse_args()
    if args.runs < 1:
        parser.error("--runs must be positive")
    prefix = ["wsl.exe", args.probe_bin] if args.wsl else [args.probe_bin]
    input_arg = wsl_path(args.input) if args.wsl else str(args.input)
    output_arg = wsl_path(args.output_dir) if args.wsl else str(args.output_dir)
    report_arg = wsl_path(args.report) if args.wsl else str(args.report)
    command = prefix + ["--input", input_arg, "--output-dir", output_arg, "--report", report_arg]
    subprocess.run(command, check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    totals, library, patches = [], [], []
    for _ in range(args.runs):
        subprocess.run(command, check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        data = json.loads(args.report.read_text(encoding="utf-8"))
        totals.append(data["elapsed_ms"])
        library.append(data["umya_library_elapsed_ms"])
        patches.append(data["ooxml_patch_elapsed_ms"])
    summarize = lambda values: {"median": statistics.median(values), "min": min(values), "max": max(values)}
    result = {
        "warmup_runs": 1,
        "measured_runs": args.runs,
        "combined_internal_ms": summarize(totals),
        "umya_library_internal_ms": summarize(library),
        "ooxml_patch_internal_ms": summarize(patches),
        "report": str(args.report),
    }
    serialized = json.dumps(result, ensure_ascii=False, indent=2)
    args.report.with_name("umya_performance.json").write_text(serialized + "\n", encoding="utf-8")
    print(serialized)


if __name__ == "__main__":
    main()
