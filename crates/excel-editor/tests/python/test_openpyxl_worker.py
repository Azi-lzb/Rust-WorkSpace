import hashlib
import json
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest
import zipfile
from xml.etree import ElementTree as ET

import openpyxl
from openpyxl.styles import Font

WORKER = Path(__file__).resolve().parents[2] / "python" / "openpyxl_worker.py"


def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def run_worker(plan, output):
    request = output.with_suffix(".request.json")
    response = output.with_suffix(".response.json")
    request.write_text(json.dumps({
        "schema_version": 1,
        "request_id": plan["request_id"],
        "format": "xlsx",
        "input": plan["input"],
        "output": str(output),
        "plan": plan,
    }), encoding="utf-8")
    process = subprocess.run(
        [sys.executable, "-I", str(WORKER), "--request", str(request), "--response", str(response)],
        capture_output=True,
        text=True,
        timeout=30,
    )
    return process, json.loads(response.read_text(encoding="utf-8"))


def operation(kind, **values):
    return {"operation": kind, **values}


class OpenpyxlWorkerTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(prefix="openpyxl-worker-test-")
        self.root = Path(self.temp.name)
        self.source = self.root / "source.xlsx"
        book = openpyxl.Workbook()
        sheet = book.active
        sheet.title = "Data"
        sheet["A1"] = "before"
        sheet["B1"] = 3
        sheet["C2"] = '=A1&"!"'
        sheet.merge_cells("D1:E1")
        sheet["D1"] = "merged"
        sheet["A1"].comment = openpyxl.comments.Comment("old note", "tester")
        sheet["B2"].comment = openpyxl.comments.Comment("retain note", "tester")
        book.create_sheet("Notes")["A1"] = "keep"
        book.create_sheet("DeleteMe")
        book.save(self.source)
        self.import_source = self.root / "import.xlsx"
        imported = openpyxl.Workbook()
        src = imported.active
        src.title = "Source"
        src["A1"] = "imported"
        src["B1"] = 9
        src["A1"].font = Font(bold=True)
        src["A1"].comment = openpyxl.comments.Comment("copied note", "source-author")
        src.merge_cells("A2:B2")
        src["A2"] = "merged import"
        src.row_dimensions[1].height = 25
        src.column_dimensions["A"].width = 22
        imported.save(self.import_source)

    def tearDown(self):
        self.temp.cleanup()

    def base_plan(self, output, operations):
        return {
            "request_id": "worker-test-1",
            "input": str(self.source),
            "output": str(output),
            "format": "xlsx",
            "workbook": {"format": "xlsx", "inventory_complete": True, "features_present": []},
            "operations": operations,
            "preserve": [],
        }

    def test_batch_modifies_one_workbook_and_preserves_supported_structures(self):
        output = self.root / "candidate.xlsx"
        source_hash = digest(self.source)
        import_hash = digest(self.import_source)
        operations = [
            operation("set_cell_value", sheet="Data", cell="A1", value={"type": "decimal", "value": "12.75"}),
            operation("set_formula_text", sheet="Data", cell="B3", formula="=A1*2"),
            operation("set_comment", sheet="Data", cell="A1", text="new note", author="worker"),
            operation("set_column_width", sheet="Data", column_index=2, width=18.5),
            operation("set_row_height", sheet="Data", row_index=3, height=27.0),
            operation("set_cell_fill_color", sheet="Data", cell="A1", argb="FF112233"),
            operation("set_page_settings", sheet="Data", settings={
                "orientation": "landscape", "fit_to_page": True, "fit_to_width": 1,
                "fit_to_height": 0, "clear_scale": True, "paper_size": 9,
                "print_area": "A1:E3", "horizontal_centered": True,
                "vertical_centered": False, "footer": "&P",
            }),
            operation("copy_range", source_sheet="Data", source_range="A1:B1",
                      target_sheet="Data", target_start="H5", copy_values=True,
                      copy_formulas=False, copy_styles=True, copy_comments=True),
            operation("import_range", source_workbook=str(self.import_source), source_sheet="Source",
                      source_range="A1:B2", target_sheet="Data", target_start="J1",
                      copy_values=True, copy_formulas=True, style_mode="full", copy_comments=True,
                      copy_merges=True, copy_row_heights=True, copy_column_widths=True,
                      copy_row_hidden=False, copy_column_hidden=False, copy_outline_levels=False,
                      copy_all_row_dimensions=False, copy_all_column_dimensions=False,
                      materialize_blank_cells=True, hide_zero=False),
            operation("rename_sheet", old_name="Notes", new_name="Information"),
            operation("create_sheet", name="Generated"),
            operation("delete_sheet", name="DeleteMe"),
        ]
        plan = self.base_plan(output, operations)
        process, response = run_worker(plan, output)
        self.assertEqual(process.returncode, 0, process.stderr)
        self.assertEqual(response["status"], "applied", response)
        self.assertEqual(response["operation_count"], len(operations))
        self.assertEqual(response["input_sha256_before"], response["input_sha256_after"])
        self.assertEqual(digest(self.source), source_hash)
        self.assertEqual(digest(self.import_source), import_hash)

        saved = openpyxl.load_workbook(output, data_only=False)
        self.assertEqual(saved.sheetnames, ["Data", "Information", "Generated"])
        sheet = saved["Data"]
        self.assertEqual(sheet["A1"].value, 12.75)
        self.assertEqual(sheet["B3"].value, "=A1*2")
        self.assertEqual(sheet["C2"].value, '=A1&"!"')
        self.assertEqual(sheet["A1"].comment.text, "new note")
        self.assertEqual(sheet["B2"].comment.text, "retain note")
        self.assertEqual(sheet["A1"].fill.fgColor.rgb, "FF112233")
        self.assertAlmostEqual(sheet.column_dimensions["B"].width, 18.5)
        self.assertAlmostEqual(sheet.row_dimensions[3].height, 27.0)
        self.assertEqual(sheet.print_area, "'Data'!$A$1:$E$3")
        self.assertEqual(sheet.page_setup.orientation, "landscape")
        self.assertTrue(sheet.sheet_properties.pageSetUpPr.fitToPage)
        self.assertEqual(sheet.oddFooter.center.text, "&P")
        self.assertEqual(sheet["H5"].value, 12.75)
        self.assertEqual(sheet["I5"].value, 3)
        self.assertEqual(sheet["H5"].comment.text, "new note")
        self.assertEqual(sheet["J1"].value, "imported")
        self.assertTrue(sheet["J1"].font.bold)
        self.assertEqual(sheet["J1"].comment.text, "copied note")
        self.assertIn("J2:K2", {item.coord for item in sheet.merged_cells.ranges})
        self.assertAlmostEqual(sheet.row_dimensions[1].height, 25)
        self.assertAlmostEqual(sheet.column_dimensions["J"].width, 22)
        saved.close()

    def test_unsupported_row_shift_is_rejected_before_output_write(self):
        output = self.root / "unsupported.xlsx"
        plan = self.base_plan(output, [operation("insert_rows", sheet="Data", first_row=2, count=1)])
        process, response = run_worker(plan, output)
        self.assertEqual(process.returncode, 0)
        self.assertEqual(response["status"], "unsupported_before_write")
        self.assertEqual(response["error"]["kind"], "unsupported_operation")
        self.assertFalse(output.exists())

    def test_formula_cache_is_rejected_before_write(self):
        cached = self.root / "cached.xlsx"
        with zipfile.ZipFile(self.source, "r") as original, zipfile.ZipFile(cached, "w", zipfile.ZIP_DEFLATED) as patched:
            for item in original.infolist():
                data = original.read(item.filename)
                if item.filename == "xl/worksheets/sheet1.xml":
                    root = ET.fromstring(data)
                    main_ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                    cell = root.find(f".//{{{main_ns}}}c[@r='C2']")
                    self.assertIsNotNone(cell)
                    cached_value = cell.find(f"{{{main_ns}}}v")
                    self.assertIsNotNone(cached_value)
                    cached_value.text = "1"
                    ET.register_namespace("", main_ns)
                    data = ET.tostring(root, encoding="utf-8", xml_declaration=True)
                patched.writestr(item, data)
        self.assertTrue(zipfile.is_zipfile(cached))
        self.source = cached
        output = self.root / "cached-output.xlsx"
        plan = self.base_plan(output, [operation("set_cell_value", sheet="Data", cell="A1", value={"type": "text", "value": "changed"})])
        process, response = run_worker(plan, output)
        self.assertEqual(process.returncode, 0)
        self.assertEqual(response["status"], "unsupported_before_write")
        self.assertEqual(response["error"]["kind"], "formula_cache_not_preserved")
        self.assertFalse(output.exists())

    def test_operation_error_is_structured_and_does_not_publish_output(self):
        output = self.root / "failed.xlsx"
        plan = self.base_plan(output, [operation("set_cell_value", sheet="Missing", cell="A1", value={"type": "text", "value": "x"})])
        process, response = run_worker(plan, output)
        self.assertNotEqual(process.returncode, 0)
        self.assertEqual(response["status"], "failed")
        self.assertEqual(response["error"]["kind"], "operation_failed")
        self.assertFalse(output.exists())


if __name__ == "__main__":
    unittest.main()
