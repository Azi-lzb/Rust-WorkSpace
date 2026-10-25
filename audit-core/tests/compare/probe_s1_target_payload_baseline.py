#!/usr/bin/env python3
"""Probe whether the Python S1 formula-copy baseline preserves merged non-anchor payloads.

The shared openpyxl worker preflight rejects target workbooks that store content
in merged non-anchor cells, because a plain openpyxl load+save drops that
content. This probe answers, per real audit workbook:

1. how many such payload cells exist (same semantics as the shared guard),
2. how many the Python baseline (``copy_formula_ranges_from_template``, which
   is itself an openpyxl load + modify + save cycle) retains,
3. how many a plain openpyxl load+save round trip retains.

Every input is copied to a temporary directory before any engine touches it.
The report persists aggregate counts only: no path, workbook name, sheet name,
cell coordinate, value or formula text is written.
"""

from __future__ import annotations

import argparse
import json
import sys
import tempfile
import time
import zipfile
from pathlib import Path
from xml.etree import ElementTree

from openpyxl.utils.cell import coordinate_to_tuple, range_boundaries


ROOT = Path(__file__).resolve().parents[3]
MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
OFFICE_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def sheet_name_map(archive: zipfile.ZipFile) -> dict[str, str]:
    """Map worksheet part names (xl/worksheets/sheetN.xml) to sheet titles."""
    workbook_root = ElementTree.fromstring(archive.read("xl/workbook.xml"))
    relations_root = ElementTree.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    relation_targets = {
        item.get("Id"): item.get("Target")
        for item in relations_root.findall(f"{{{REL_NS}}}Relationship")
    }
    mapping: dict[str, str] = {}
    for sheet in workbook_root.findall(f".//{{{MAIN_NS}}}sheet"):
        relation_id = sheet.get(f"{{{OFFICE_REL_NS}}}id")
        target = relation_targets.get(relation_id)
        if not target:
            continue
        part = target.lstrip("/")
        if not part.startswith("xl/"):
            part = f"xl/{part}"
        mapping[part] = sheet.get("name")
    return mapping


def merged_bounds(archive: zipfile.ZipFile, part: str) -> list[tuple[int, int, int, int]]:
    root = ElementTree.fromstring(archive.read(part))
    bounds = []
    for node in root.findall(f".//{{{MAIN_NS}}}mergeCell"):
        reference = node.get("ref")
        if not reference:
            continue
        min_col, min_row, max_col, max_row = range_boundaries(reference.replace("$", ""))
        bounds.append((min_row, min_col, max_row, max_col))
    return bounds


def cell_text(element) -> str:
    return element.text or ""


def payload_content(cell) -> tuple | None:
    """Mirror the shared guard's stored-content semantics for one <c> element."""
    content = []
    for child in cell:
        if child.tag == f"{{{MAIN_NS}}}f":
            content.append(("f", child.get("t"), cell_text(child)))
        elif child.tag == f"{{{MAIN_NS}}}v":
            if cell_text(child) != "":
                content.append(("v", None, cell_text(child)))
        elif child.tag == f"{{{MAIN_NS}}}is":
            inline = "".join(
                node.text or ""
                for node in child.findall(f".//{{{MAIN_NS}}}t")
            )
            if inline != "":
                content.append(("is", None, inline))
    if not content:
        return None
    return tuple(content)


def payload_census(archive: zipfile.ZipFile) -> dict:
    """Collect merged non-anchor payload cells per sheet (guard semantics)."""
    name_map = sheet_name_map(archive)
    payloads: dict[tuple[str, str], tuple] = {}
    merged_total = 0
    payload_bearing_areas = 0
    for part, sheet_name in sorted(name_map.items()):
        if "/_rels/" in part or not part.startswith("xl/worksheets/"):
            continue
        root = ElementTree.fromstring(archive.read(part))
        bounds = merged_bounds(archive, part)
        merged_total += len(bounds)
        areas_with_payload = 0
        for index, (min_row, min_col, max_row, max_col) in enumerate(bounds):
            has_payload_in_area = False
            for cell in root.findall(f".//{{{MAIN_NS}}}c"):
                address = cell.get("r")
                if not address:
                    continue
                row, column = coordinate_to_tuple(address)
                if not (
                    min_row <= row <= max_row
                    and min_col <= column <= max_col
                ):
                    continue
                if row == min_row and column == min_col:
                    continue
                content = payload_content(cell)
                if content is not None:
                    payloads[(sheet_name, address.replace("$", ""))] = content
                    has_payload_in_area = True
            if has_payload_in_area:
                areas_with_payload += 1
        payload_bearing_areas += areas_with_payload
    return {
        "payloads": payloads,
        "merged_range_count": merged_total,
        "payload_bearing_merged_areas": payload_bearing_areas,
    }


def inside_copy_ranges(
    sheet_name: str, coordinate: str, ranges: list[dict[str, str]]
) -> bool:
    from openpyxl.utils.cell import coordinate_to_tuple

    row, column = coordinate_to_tuple(coordinate)
    for item in ranges:
        if item["sheet_name"] != sheet_name:
            continue
        min_col, min_row, max_col, max_row = range_boundaries(
            item["address"].replace("$", "")
        )
        if min_row <= row <= max_row and min_col <= column <= max_col:
            return True
    return False


def census_summary(
    archive: zipfile.ZipFile, ranges: list[dict[str, str]]
) -> dict:
    census = payload_census(archive)
    inside = 0
    for (sheet_name, coordinate) in census["payloads"]:
        if inside_copy_ranges(sheet_name, coordinate, ranges):
            inside += 1
    return {
        "payload_cells_total": len(census["payloads"]),
        "payload_cells_inside_copy_ranges": inside,
        "payload_cells_outside_copy_ranges": len(census["payloads"]) - inside,
        "payload_bearing_merged_areas": census["payload_bearing_merged_areas"],
        "merged_range_count": census["merged_range_count"],
    }


def retention(
    original_payloads: dict[tuple[str, str], tuple],
    output_payloads: dict[tuple[str, str], tuple],
) -> dict:
    retained = set(original_payloads) & set(output_payloads)
    retained_exact = sum(
        1 for key in retained if original_payloads[key] == output_payloads[key]
    )
    gained = set(output_payloads) - set(original_payloads)
    return {
        "payload_cells_total": len(output_payloads),
        "payload_cells_retained": len(retained),
        "payload_cells_retained_exact_content": retained_exact,
        "payload_cells_dropped": len(original_payloads) - len(retained),
        "payload_cells_gained": len(gained),
    }


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--source-root", type=Path, required=True)
    parser.add_argument("--template-workbook", type=Path, required=True)
    parser.add_argument("--mapping-config", type=Path)
    parser.add_argument("--input-dir", type=Path, required=True)
    parser.add_argument("--samples", type=int, default=0,
                        help="limit to the first N sorted inputs; 0 means all")
    parser.add_argument("--report", type=Path, required=True)
    args = parser.parse_args()

    source_root = args.source_root.resolve(strict=True)
    template_path = args.template_workbook.resolve(strict=True)
    input_dir = args.input_dir.resolve(strict=True)
    report_path = args.report.resolve()
    sys.path.insert(0, str(source_root / "core" / "src"))
    sys.path.insert(0, str(Path(__file__).resolve().parent))

    import compare_s1_formula_copy_candidate as driver
    from base_audit.native.openpyxl_workbook import copy_formula_ranges_from_template

    inputs = sorted(
        path
        for path in input_dir.glob("*.xlsx")
        if not path.name.startswith("~$")
    )
    if args.samples > 0:
        inputs = inputs[: args.samples]
    if not inputs:
        raise SystemExit("no candidate workbooks found in the input directory")

    with tempfile.TemporaryDirectory(prefix="audit-s1-payload-probe-") as temporary:
        directory = Path(temporary)
        mapping_path = None
        if args.mapping_config:
            mapping_path = directory / "mapping.xlsx"
            driver.shutil.copy2(args.mapping_config.resolve(strict=True), mapping_path)

        definition = driver.resolve_real_ranges(source_root, template_path, mapping_path)
        ranges = [
            {"sheet_name": item.sheet_name, "address": item.address}
            for item in definition.copy_ranges
        ]
        scope_stats = driver.range_stats(template_path, ranges)

        samples = []
        for ordinal, source in enumerate(inputs, start=1):
            started = time.perf_counter()
            original_bytes = source.read_bytes()

            original_archive = zipfile.ZipFile(source)
            original_census = census_summary(original_archive, ranges)
            original_payloads = payload_census(original_archive)["payloads"]
            original_cache = driver.formula_cache_count(source)
            original_archive.close()

            baseline_output = directory / f"baseline_{ordinal}.xlsx"
            driver.shutil.copy2(source, baseline_output)
            template_book = driver.load_workbook(
                template_path, data_only=False, keep_links=False
            )
            baseline_started = time.perf_counter()
            copy_formula_ranges_from_template(
                template_book, baseline_output, definition
            )
            baseline_ms = round((time.perf_counter() - baseline_started) * 1000)
            template_book.close()

            loadsave_input = directory / f"loadsave_{ordinal}.xlsx"
            loadsave_output = directory / f"loadsave_out_{ordinal}.xlsx"
            driver.shutil.copy2(source, loadsave_input)
            book = driver.load_workbook(loadsave_input, data_only=False, keep_links=True)
            book.save(loadsave_output)
            book.close()

            baseline_archive = zipfile.ZipFile(baseline_output)
            baseline_summary = census_summary(baseline_archive, ranges)
            baseline_payloads = payload_census(baseline_archive)["payloads"]
            baseline_cache = driver.formula_cache_count(baseline_output)
            baseline_archive.close()

            loadsave_archive = zipfile.ZipFile(loadsave_output)
            loadsave_summary = census_summary(loadsave_archive, ranges)
            loadsave_payloads = payload_census(loadsave_archive)["payloads"]
            loadsave_cache = driver.formula_cache_count(loadsave_output)
            loadsave_archive.close()

            sample = {
                "sorted_ordinal_1_based": ordinal,
                "formula_cache_nonempty_count": original_cache,
                "original": original_census,
                "python_baseline_output": {
                    **baseline_summary,
                    **retention(original_payloads, baseline_payloads),
                    "formula_cache_nonempty_count": baseline_cache,
                    "baseline_elapsed_ms": baseline_ms,
                },
                "openpyxl_load_save_output": {
                    **loadsave_summary,
                    **retention(original_payloads, loadsave_payloads),
                    "formula_cache_nonempty_count": loadsave_cache,
                },
                "input_unchanged": source.read_bytes() == original_bytes,
            }
            samples.append(sample)
            print(
                json.dumps(
                    {
                        "ordinal": ordinal,
                        "original_outside": original_census["payload_cells_outside_copy_ranges"],
                        "baseline_retained": sample["python_baseline_output"]["payload_cells_retained"],
                        "loadsave_retained": sample["openpyxl_load_save_output"]["payload_cells_retained"],
                        "elapsed_ms": round((time.perf_counter() - started) * 1000),
                    }
                ),
                flush=True,
            )

    baseline_all_dropped = all(
        sample["python_baseline_output"]["payload_cells_retained"] == 0
        for sample in samples
    )
    loadsave_all_dropped = all(
        sample["openpyxl_load_save_output"]["payload_cells_retained"] == 0
        for sample in samples
    )
    baseline_matches_loadsave = all(
        sample["python_baseline_output"]["payload_cells_retained"]
        == sample["openpyxl_load_save_output"]["payload_cells_retained"]
        for sample in samples
    )
    inputs_unchanged = all(sample["input_unchanged"] for sample in samples)

    report = {
        "schema_version": 1,
        "suite": "audit.s1_formula_copy_target_payload_baseline_probe",
        "feature_id": "S1-F01",
        "status": (
            "completed"
            if inputs_unchanged
            else "input_mutated"
        ),
        "sample_scope": {
            "kind": "real_s1_batch",
            "selected_samples": len(samples),
            "selection": "sorted xlsx inputs in the provided directory",
        },
        "template_cross_check": {
            "copy_ranges": len(ranges),
            "formula_cells": scope_stats["formula_cells"],
            "range_comment_count_python": scope_stats["range_comment_count_python"],
        },
        "payload_semantics": (
            "merged non-anchor cells with a stored formula, nonempty cached "
            "value, or nonempty inline string; identical to the shared "
            "excel-editor openpyxl preflight guard"
        ),
        "samples": samples,
        "summary": {
            "baseline_retained_zero_payloads_all_samples": baseline_all_dropped,
            "loadsave_retained_zero_payloads_all_samples": loadsave_all_dropped,
            "baseline_and_loadsave_retention_equal_all_samples": baseline_matches_loadsave,
            "inputs_unchanged": inputs_unchanged,
        },
        "data_safety": {
            "input_workbooks_modified": not inputs_unchanged,
            "outputs_written_to_temporary_directories": True,
            "file_paths_recorded": False,
            "workbook_or_sheet_names_recorded": False,
            "cell_coordinates_recorded": False,
            "business_values_or_formulas_recorded": False,
        },
    }
    report_path.parent.mkdir(parents=True, exist_ok=True)
    report_path.write_text(
        json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )
    print(
        json.dumps(
            {
                "status": report["status"],
                "samples": len(samples),
                "baseline_retained_zero": baseline_all_dropped,
                "loadsave_retained_zero": loadsave_all_dropped,
            }
        )
    )
    return 0 if inputs_unchanged else 1


if __name__ == "__main__":
    raise SystemExit(main())
