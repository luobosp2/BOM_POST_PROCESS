from dataclasses import asdict
import hashlib
import json
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

from openpyxl import Workbook

from bom_post_process.errors import BomError
from bom_post_process.validation import validate_workbook


class ValidationTests(unittest.TestCase):
    def setUp(self):
        self.temporary = TemporaryDirectory()
        self.addCleanup(self.temporary.cleanup)
        self.source = Path(self.temporary.name) / "input.xlsx"

    def write_rows(self, rows):
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "AD export"
        sheet.append(["Comment", "Description", "Designator", "Footprint", "LibRef", "Quantity"])
        for row in rows:
            sheet.append(row)
        workbook.save(self.source)
        workbook.close()

    def test_supported_aliases_and_nc_suffixes(self):
        rows = []
        for index, footprint in enumerate(["C 0402L", "C0402", "CAP_0402", "0402C", "0402"]):
            rows.append(["0402-X5R104K250/NC", "", f"C{index + 1}", footprint, "", 1])
        for index, footprint in enumerate(["RES-0402", "0402R", "R0402", "0402"]):
            rows.append([" 0402k1002(NC) ", "", f"R{index + 1}", footprint, "", 1])
        self.write_rows(rows)
        before = hashlib.sha256(self.source.read_bytes()).hexdigest()
        report = validate_workbook(self.source)
        self.assertTrue(report.ok)
        self.assertEqual((report.checked, report.skipped), (9, 0))
        self.assertEqual(report.issues, ())
        self.assertEqual(hashlib.sha256(self.source.read_bytes()).hexdigest(), before)

    def test_package_and_component_type_conflicts(self):
        self.write_rows([
            ["0402-X5R104K250", "", "C1", "C0603", "", 1],
            ["0402K1002", "", "R1", "0402C", "", 1],
            ["0603-X7R105K250", "", "C2", "R0402", "", 1],
        ])
        report = validate_workbook(self.source)
        self.assertFalse(report.ok)
        self.assertEqual((report.checked, report.errors, report.warnings), (3, 4, 0))
        self.assertEqual(
            [(issue.row, issue.code) for issue in report.issues],
            [(2, "package_mismatch"), (3, "component_type_mismatch"),
             (4, "package_mismatch"), (4, "component_type_mismatch")],
        )

    def test_unknown_models_and_ambiguous_footprints_are_unverified(self):
        self.write_rows([
            ["IC_M1206", "", "U1", "M1220", "", 1],
            ["1206B104K101NT(FH)", "", "C1", "C0402", "", 1],
            ["0402K1002", "", "R1", "R0402/C0603", "", 1],
            ["0402K1002", "", "R2", "M0402", "", 1],
        ])
        report = validate_workbook(self.source)
        self.assertTrue(report.ok)
        self.assertEqual((report.checked, report.skipped, report.warnings), (0, 4, 4))
        self.assertEqual([issue.code for issue in report.issues],
                         ["unsupported_comment", "unsupported_comment",
                          "unknown_footprint", "unknown_footprint"])

    def test_duplicate_designators_and_quantity_mismatch(self):
        self.write_rows([
            ["0402K1002", "", "R1, R1", "R0402", "", 2],
            ["0402K1002", "", "r1; R2", "R0402", "", 3],
            ["0402K1002", "", "R3", "R0402", "", 0],
            ["0402K1002", "", "", "R0402", "", 1],
        ])
        report = validate_workbook(self.source)
        self.assertEqual([issue.code for issue in report.issues],
                         ["duplicate_designator", "duplicate_designator", "quantity_mismatch",
                          "invalid_quantity", "missing_designator"])
        self.assertIn("第 2 行", report.issues[1].message)

    def test_positive_integer_quantities_and_unusual_designators(self):
        self.write_rows([
            ["0402K1002", "", "+1.0V, VIN-1", "R0402", "", "2.0"],
            ["0402K1002", "", "R2，R3、R4\nR5", "R0402", "", 4.0],
            ["0402K1002", "", "R6 R7\tR8", "R0402", "", 3],
        ])
        report = validate_workbook(self.source)
        self.assertEqual(report.issues, ())

    def test_missing_values_are_errors(self):
        self.write_rows([
            [None, "", "C1", "C0402", "", 1],
            ["0402K1002", "", "R1", None, "", 1],
            ["0402K1002", "", "R2", "R0402", "", "NaN"],
            ["0402K1002", "", "R3", "R0402", "", 1.5],
        ])
        report = validate_workbook(self.source)
        self.assertFalse(report.ok)
        self.assertEqual([issue.code for issue in report.issues],
                         ["missing_comment", "missing_footprint", "invalid_quantity", "invalid_quantity"])

    def test_report_is_json_serializable_with_statistics(self):
        self.write_rows([["0402K1002", "", "R1", "R0603", "", 1]])
        report = validate_workbook(self.source, sheet="AD export")
        result = json.loads(json.dumps(asdict(report), ensure_ascii=False))
        self.assertEqual(result["sheet"], "AD export")
        self.assertEqual(result["errors"], 1)
        self.assertEqual(result["warnings"], 0)
        self.assertFalse(result["ok"])
        self.assertEqual(result["issues"][0]["designator"], "R1")

    def test_empty_table_is_not_reported_as_valid(self):
        self.write_rows([])
        with self.assertRaisesRegex(BomError, "没有 BOM 数据行"):
            validate_workbook(self.source)

    def test_supplied_workbook(self):
        source = Path(__file__).resolve().parents[1] / "input" / "DAQ_1010A_V1R0.xlsx"
        before = hashlib.sha256(source.read_bytes()).hexdigest()
        report = validate_workbook(source)
        self.assertEqual((report.checked, report.skipped, report.errors, report.warnings),
                         (91, 51, 0, 51))
        self.assertEqual({issue.code for issue in report.issues}, {"unsupported_comment"})
        self.assertEqual(hashlib.sha256(source.read_bytes()).hexdigest(), before)


if __name__ == "__main__":
    unittest.main()
