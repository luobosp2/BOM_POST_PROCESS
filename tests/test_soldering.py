from copy import copy
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils.datetime import CALENDAR_MAC_1904

from bom_post_process.errors import BomError
from bom_post_process.soldering import gen_soldering
from bom_post_process.values import gen_value
from bom_post_process.workbook import AD_HEADERS, BOM_HEADERS, header_map, split_designators


REVIEW_HEADERS = (*AD_HEADERS, "name", "note", "not_bought", "not_welded")


class SolderingTests(unittest.TestCase):
    def setUp(self):
        self.temporary = TemporaryDirectory()
        self.addCleanup(self.temporary.cleanup)
        self.folder = Path(self.temporary.name)
        self.source = self.folder / "样例_已解析.xlsx"
        self.output = self.folder / "焊接.xlsx"
        book = Workbook()
        sheet = book.active
        sheet.title = "已确认"
        sheet.append(REVIEW_HEADERS)
        for row in [
            ["0402-X5R104K250", "原始描述", "C1,C2", "C0402", "Cap", 2,
             "人工电容名称", "  人工改为100nF/50V\n指定供应商  ", None, None],
            ["0402K1003", "", "R1", "R0402", "Res", 1,
             "人工电阻名称", "人工修改：100k/0.1%", "不买但要焊", None],
            ["0402K1003/NC", "", "R2,R3", "R0402", "Res", 2,
             "电阻", "手工决定仍焊接", None, None],
            ["MARK", "", "M1", "MARK", "MarkPoint", 1,
             "定位点", "仅板上图形", None, "不焊接"],
            ["TP", "", "TP1", "TP", "TestPoint", 1,
             "测试点", "裸露焊盘", "不买", "不焊接"],
            ["PCB制板", "", "", "", "PCB", 1,
             "PCB制板", "用户自行决定是否放在焊接表", None, None],
        ]:
            sheet.append(row)
        sheet["H2"].font = Font(name="宋体", size=12, bold=True)
        sheet["H2"].fill = PatternFill("solid", fgColor="FFFF00")
        sheet["H2"].alignment = Alignment(wrap_text=True)
        sheet.row_dimensions[2].height = 48
        sheet.column_dimensions["H"].width = 52
        sheet.column_dimensions["G"].width = 24
        sheet.freeze_panes = "A2"
        book.save(self.source)
        book.close()

    def edit(self, action, destination=None):
        book = load_workbook(self.source)
        action(book)
        path = destination or self.source
        book.save(path)
        book.close()
        return path

    def test_only_welding_flag_filters_and_manual_values_styles_survive(self):
        before = self.source.read_bytes()
        report = gen_soldering(self.source, self.output)
        self.assertEqual(self.source.read_bytes(), before)
        self.assertEqual((report.written, report.excluded, report.pcb_placeholder), (4, 2, 0))
        source, result = load_workbook(self.source), load_workbook(self.output)
        try:
            self.assertEqual(result.sheetnames, ["焊接表"])
            target = result.active
            self.assertEqual(tuple(cell.value for cell in target[1]), BOM_HEADERS)
            for output_row, input_row in enumerate([2, 3, 4, 7], start=2):
                self.assertEqual(target.cell(output_row, 1).value, output_row - 1)
                for output_column, input_column in enumerate([1, 7, 6, 3, 8], start=2):
                    original = source.active.cell(input_row, input_column)
                    copied = target.cell(output_row, output_column)
                    self.assertEqual(original.value, copied.value)
                    for attribute in ("font", "fill", "alignment", "border", "number_format"):
                        self.assertEqual(copy(getattr(original, attribute)), copy(getattr(copied, attribute)))
            self.assertEqual(target.column_dimensions["F"].width, 52)
            self.assertEqual(target.column_dimensions["C"].width, 24)
            self.assertEqual(target.row_dimensions[2].height, 48)
            self.assertEqual(target["B4"].value, "0402K1003/NC")
            self.assertEqual(target["B5"].value, "PCB制板")
        finally:
            source.close()
            result.close()

    def test_marker_zero_false_and_text_exclude_while_whitespace_is_blank(self):
        for index, (value, excluded) in enumerate([(None, 2), ("", 2), (" \t ", 2),
                                                  (0, 3), (False, 3), ("False", 3), ("否", 3)]):
            with self.subTest(value=value):
                source = self.edit(lambda book: setattr(book.active["J2"], "value", value),
                                   self.folder / f"marker{index}.xlsx")
                report = gen_soldering(source, self.folder / f"marker_out{index}.xlsx")
                self.assertEqual(report.excluded, excluded)

    def test_purchase_flag_is_ignored_even_if_formula_or_error(self):
        for index, value in enumerate(("=1+1", "#REF!")):
            source = self.edit(lambda book: setattr(book.active["I2"], "value", value),
                               self.folder / f"ignored{index}.xlsx")
            report = gen_soldering(source, self.folder / f"ignored_out{index}.xlsx")
            self.assertEqual(report.written, 4)

    def test_blank_name_is_not_inferred_and_literal_formula_text_survives(self):
        def change(book):
            book.active["G2"] = None
            book.active["H2"] = "=manual literal"
            book.active["H2"].data_type = "s"
        self.edit(change)
        gen_soldering(self.source, self.output)
        book = load_workbook(self.output)
        try:
            self.assertIsNone(book.active["C2"].value)
            self.assertEqual(book.active["F2"].value, "=manual literal")
            self.assertEqual(book.active["F2"].data_type, "s")
        finally:
            book.close()

    def test_missing_review_columns_requires_gen_value(self):
        for header in ("name", "note", "not_bought", "not_welded"):
            with self.subTest(header=header):
                source = self.edit(lambda book: book.active.delete_cols(header_map(book.active)[header]),
                                   self.folder / f"missing_{header}.xlsx")
                with self.assertRaisesRegex(BomError, "先运行 gen_value"):
                    gen_soldering(source, self.output)
                self.assertFalse(self.output.exists())

    def test_explicit_sheet_selects_reviewed_table_and_preserves_date_epoch(self):
        def change(book):
            book.copy_worksheet(book.active).title = "另一份"
            book.epoch = CALENDAR_MAC_1904
        self.edit(change)
        with self.assertRaisesRegex(BomError, "多个已解析"):
            gen_soldering(self.source, self.output)
        report = gen_soldering(self.source, self.output, sheet="另一份")
        self.assertEqual(report.source_sheet, "另一份")
        result = load_workbook(self.output)
        try:
            self.assertEqual(result.epoch, CALENDAR_MAC_1904)
        finally:
            result.close()

    def test_invalid_selected_data_and_layout_rejected_without_output(self):
        for mode in ("name", "note", "not_welded", "comment", "merged", "empty"):
            with self.subTest(mode=mode):
                def change(book):
                    sheet = book.active
                    if mode == "merged":
                        sheet.merge_cells("G2:H2")
                    elif mode == "empty":
                        sheet.delete_rows(2, 20)
                    else:
                        sheet.cell(2, header_map(sheet)[mode], "=1+1")
                source = self.edit(change, self.folder / f"invalid_{mode}.xlsx")
                with self.assertRaises(BomError):
                    gen_soldering(source, self.output)
                self.assertFalse(self.output.exists())

    def test_filtered_rows_are_not_exported_or_validated(self):
        def change(book):
            book.active["H5"] = "=1+1"
            for row in range(2, book.active.max_row + 1):
                book.active.cell(row, 10, "排除")
        self.edit(change)
        report = gen_soldering(self.source)
        self.assertEqual(report.output.name, "样例_焊接表.xlsx")
        self.assertEqual((report.written, report.excluded), (0, 6))
        book = load_workbook(report.output)
        try:
            self.assertEqual(book.sheetnames, ["焊接表"])
            self.assertEqual(book.active.max_row, 1)
        finally:
            book.close()

    def test_output_and_source_never_overwritten(self):
        self.output.write_bytes(b"keep existing")
        for output in (self.source, self.output):
            before = output.read_bytes()
            with self.assertRaises(BomError):
                gen_soldering(self.source, output)
            self.assertEqual(output.read_bytes(), before)

    def test_real_input_manual_name_and_note_survive_after_nc_defaults(self):
        source = Path(__file__).resolve().parents[1] / "input" / "DAQ_1010A_V1R0.xlsx"
        parsed = self.folder / "parsed.xlsx"
        gen_value(source, parsed)
        book = load_workbook(parsed)
        headers = header_map(book.active)
        book.active.cell(63, headers["name"], "精密电阻")
        book.active.cell(63, headers["note"], "人工确认100k/0.1%，勿重新解析")
        book.save(parsed)
        book.close()
        report = gen_soldering(parsed, self.output)
        self.assertEqual((report.written, report.excluded), (134, 8))
        book = load_workbook(self.output)
        try:
            reviewed = [row for row in list(book.active.values)[1:] if "R1" in split_designators(row[4])]
            self.assertEqual(reviewed[0][2], "精密电阻")
            self.assertEqual(reviewed[0][5], "人工确认100k/0.1%，勿重新解析")
        finally:
            book.close()


if __name__ == "__main__":
    unittest.main()
