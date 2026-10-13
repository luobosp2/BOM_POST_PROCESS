from contextlib import redirect_stderr, redirect_stdout
from io import StringIO
import json
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import patch

from openpyxl import Workbook, load_workbook

from bom_post_process.cli import main


class CliTests(unittest.TestCase):
    def setUp(self):
        self.temporary = TemporaryDirectory()
        self.addCleanup(self.temporary.cleanup)
        self.folder = Path(self.temporary.name)
        self.source = self.folder / "input.xlsx"
        workbook = Workbook()
        workbook.active.append(["Comment", "Description", "Designator", "Footprint", "LibRef", "Quantity"])
        workbook.active.append(["0402K1002", "", "R1", "R0402", "Res", 1])
        workbook.save(self.source)
        workbook.close()

    def run_cli(self, *args):
        stdout, stderr = StringIO(), StringIO()
        with redirect_stdout(stdout), redirect_stderr(stderr):
            code = main([str(arg) for arg in args])
        return code, stdout.getvalue(), stderr.getvalue()

    def test_validate_reports_conflict_and_preserves_input(self):
        workbook = load_workbook(self.source)
        workbook.active["D2"] = "R0603"
        workbook.save(self.source)
        workbook.close()
        before = self.source.read_bytes()
        report = self.folder / "report.json"
        code, _, stderr = self.run_cli("validate", self.source, "-o", report)
        self.assertEqual(code, 1)
        self.assertIn("0603", stderr)
        data = json.loads(report.read_text(encoding="utf-8"))
        self.assertFalse(data["ok"])
        self.assertEqual(data["errors"], 1)
        self.assertEqual(self.source.read_bytes(), before)

    def test_gen_value_and_collision(self):
        output = self.folder / "nested" / "parsed.xlsx"
        code, _, _ = self.run_cli("gen_value", self.source, "-o", output)
        self.assertEqual(code, 0)
        workbook = load_workbook(output)
        self.assertEqual(workbook.active["G2"].value, "10k/0402/1%")
        self.assertEqual(len(workbook.worksheets), 1)
        self.assertNotIn("焊接表", workbook.sheetnames)
        workbook.close()
        before = output.read_bytes()
        code, _, stderr = self.run_cli("gen_value", self.source, "-o", output)
        self.assertEqual(code, 1)
        self.assertIn("已存在", stderr)
        self.assertEqual(output.read_bytes(), before)

    def test_all_is_not_a_command(self):
        with self.assertRaises(SystemExit) as caught:
            self.run_cli("all", self.source)
        self.assertEqual(caught.exception.code, 2)

    def test_manual_selection_survives_soldering_export(self):
        parsed = self.folder / "parsed.xlsx"
        soldering = self.folder / "soldering.xlsx"
        code, _, _ = self.run_cli("gen_value", self.source, "-o", parsed)
        self.assertEqual(code, 0)
        workbook = load_workbook(parsed)
        note = "10k/0402/0.1%\n选型：低温漂，指定厂家 & 型号"
        workbook.active["G2"] = note
        workbook.save(parsed)
        workbook.close()
        code, _, _ = self.run_cli("gen_soldering", parsed, "-o", soldering)
        self.assertEqual(code, 0)
        workbook = load_workbook(soldering)
        self.assertEqual(workbook["焊接表"]["F2"].value, note)
        workbook.close()

    def test_standard_export_is_an_independent_command(self):
        from bom_post_process.standard import StandardReport

        word = self.folder / "result.docx"
        report = StandardReport(word, word.with_suffix(".pdf"), word.with_suffix(".json"), 1, 1, "焊接表")
        with patch("bom_post_process.standard.export_standard_bom", return_value=report) as export:
            code, _, _ = self.run_cli("gen_bom_standard", self.source, "-o", word)
        self.assertEqual(code, 0)
        export.assert_called_once_with(self.source, word, sheet=None, template=None, config=None)

    def test_missing_file_is_a_friendly_error(self):
        code, _, stderr = self.run_cli("validate", self.folder / "missing.xlsx")
        self.assertEqual(code, 1)
        self.assertIn("不存在", stderr)
        self.assertNotIn("Traceback", stderr)


if __name__ == "__main__":
    unittest.main()
