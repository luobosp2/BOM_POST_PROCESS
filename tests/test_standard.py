"""数据映射、模板结构和实际交付结果的回归检查；无第三方依赖。"""
from pathlib import Path
import json
import tempfile
import unittest
from unittest.mock import Mock, patch
import xml.etree.ElementTree as ET
import zipfile

from bom_post_process.standard import (COLUMNS, NS, RESOURCES, StandardError, Template, W,
                                      _publish, category_label, category_starts, export_standard_bom,
                                      material_category, paginate, plain_text, read_bom,
                                      read_exported_content)


def workbook(path, sheets):
    s = "{" + NS["s"] + "}"
    root = ET.Element(s + "workbook")
    sheet_list = ET.SubElement(root, s + "sheets")
    relations = ET.Element("Relationships")
    with zipfile.ZipFile(path, "w") as archive:
        for index, (name, rows) in enumerate(sheets, 1):
            ET.SubElement(sheet_list, s + "sheet", {"name": name, "{" + NS["r"] + "}id": f"rId{index}"})
            ET.SubElement(relations, "Relationship", Id=f"rId{index}", Target=f"worksheets/sheet{index}.xml")
            sheet = ET.Element(s + "worksheet")
            data = ET.SubElement(sheet, s + "sheetData")
            for ri, values in enumerate(rows, 1):
                row = ET.SubElement(data, s + "row", r=str(ri))
                for ci, text in enumerate(values):
                    if text is None:
                        continue
                    cell = ET.SubElement(row, s + "c", r=f"{chr(65 + ci)}{ri}", t="inlineStr")
                    if isinstance(text, dict):
                        cell.attrib.pop("t")
                        ET.SubElement(cell, s + "f").text = text["formula"]
                        if "value" in text:
                            ET.SubElement(cell, s + "v").text = text["value"]
                    else:
                        ET.SubElement(ET.SubElement(cell, s + "is"), s + "t").text = text
            archive.writestr(f"xl/worksheets/sheet{index}.xml", ET.tostring(sheet))
        archive.writestr("xl/workbook.xml", ET.tostring(root))
        archive.writestr("xl/_rels/workbook.xml.rels", ET.tostring(relations))


class ExcelTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.path = Path(self.temp.name) / "test.xlsx"

    def test_columns_can_be_reordered_and_extra_columns_ignored(self):
        workbook(self.path, [("BOM", [["备注", "名称", "序号", "项目代号", "数量", "规格型号/代号", "辅助"],
                                      ["中文\n第二行", "电阻", "001", "R1, R2", "2", "0402", "忽略"],
                                      [None] * 6 + ["不应成为物料行"]])])
        self.assertEqual(read_bom(self.path)[1], [["001", "0402", "电阻", "2", "R1, R2", "中文\n第二行"]])

    def test_selected_sheet(self):
        workbook(self.path, [("说明", [["不是BOM"]]),
                             ("物料", [COLUMNS, ["1", "ABC", "器件", "1", "U1", ""]])])
        self.assertEqual(read_bom(self.path, "物料")[0], "物料")
        with self.assertRaisesRegex(ValueError, "找不到工作表"):
            read_bom(self.path, "不存在")

    def test_prefers_procurement_sheet_and_requires_unambiguous_fallback(self):
        rows = [COLUMNS, ["1", "ABC", "器件", "1", "U1", ""]]
        workbook(self.path, [("说明", [["说明文字"]]), ("已解析BOM", rows)])
        self.assertEqual(read_bom(self.path)[0], "已解析BOM")
        workbook(self.path, [("其他BOM", rows), ("焊接表", rows), ("BOM", rows)])
        self.assertEqual(read_bom(self.path)[0], "BOM")
        workbook(self.path, [("其他BOM", rows), ("焊接表", rows), ("未焊接元件", rows)])
        self.assertEqual(read_bom(self.path)[0], "其他BOM")
        workbook(self.path, [("BOM1", rows), ("BOM2", rows)])
        with self.assertRaisesRegex(StandardError, "多个工作表"):
            read_bom(self.path)

    def test_soldering_only_workbooks_require_procurement_export(self):
        rows = [COLUMNS, ["1", "ABC", "器件", "1", "U1", ""]]
        for name in ("焊接表", "未焊接元件"):
            with self.subTest(sheet=name):
                workbook(self.path, [(name, rows)])
                with self.assertRaisesRegex(StandardError, "不能作为采购 BOM.*gen_bom"):
                    read_bom(self.path)
                workbook(self.path, [(name, rows), ("BOM", rows)])
                with self.assertRaisesRegex(StandardError, "不能作为采购 BOM.*gen_bom"):
                    read_bom(self.path, name)

    def test_duplicate_required_header_rejected(self):
        workbook(self.path, [("BOM", [list(COLUMNS) + ["备注"], ["1", "R", "电阻", "1", "R1", "", ""]])])
        with self.assertRaisesRegex(StandardError, "列名重复"):
            read_bom(self.path)

    def test_missing_header_and_empty_data_rejected(self):
        workbook(self.path, [("BOM", [["名称"], ["电阻"]])])
        with self.assertRaisesRegex(ValueError, "缺少表头.*gen_bom"):
            read_bom(self.path)
        workbook(self.path, [("BOM", [COLUMNS])])
        with self.assertRaisesRegex(ValueError, "没有 BOM 数据"):
            read_bom(self.path)

    def test_formula_requires_cached_result(self):
        workbook(self.path, [("BOM", [COLUMNS, ["1", "R", "电阻", {"formula": "1+1"}, "R1, R2", ""]])])
        with self.assertRaisesRegex(ValueError, "没有计算结果"):
            read_bom(self.path)
        workbook(self.path, [("BOM", [COLUMNS, ["1", "R", "电阻", {"formula": "1+1", "value": "2"}, "R1, R2", ""]])])
        self.assertEqual(read_bom(self.path)[1][0][3], "2")


class TemplateTests(unittest.TestCase):
    def test_first_and_continuation_footer_and_data_are_distinct(self):
        template = Template(RESOURCES / "bom_template.docx", {"公司名": "测试公司"})
        values = ["138", "SFP+", "插座", "1", "XS3", "第一行\n第二行 & <>"]
        first = template.table([values], True, 1, 2)
        following = template.table([values], False, 2, 2)
        self.assertIn("标准化", plain_text(first))
        self.assertNotIn("标准化", plain_text(following))
        self.assertIn("共 2 张    第 1 张", plain_text(first))
        self.assertIn("共 2 张    第 2 张", plain_text(following))
        self.assertIn("测试公司", plain_text(first))
        for table in (first, following):
            for row in table.findall("w:tr", NS):
                self.assertIsNotNone(row.find("w:trPr/w:cantSplit", NS))
        row = next(row for row in first.findall("w:tr", NS)
                   if len(row.findall("w:tc", NS)) == 9
                   and plain_text(row.findall("w:tc", NS)[1]) == "138")
        self.assertEqual([plain_text(c) for c in row.findall("w:tc", NS)[1:6]], values[:5])
        self.assertEqual(len(row.findall(".//w:br", NS)), 1)

    def test_categories_follow_materials_without_changing_their_order(self):
        template = Template(RESOURCES / "bom_template.docx")
        records = [["01", " PCB制板 ", "PCB", "5", "", ""],
                   ["02", "BOARD2", "PCB制板", "2", "", ""],
                   ["03", "0402", "电阻", "1", "R1", "两行\n备注"],
                   ["04", "PART", "外购件", "1", "U1", ""],
                   ["05", "PCB制板", "PCB", "1", "", ""]]
        # 同类物料跨页不重加标题，新分类在新页紧跟表头，来回切换不排序。
        groups = [[records[0]], [records[1]], records[2:4], [records[4]]]
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "categories.docx"
            template.write(path, groups)
            exported, starts = read_exported_content(path, template)
            self.assertEqual(exported, records)
            self.assertEqual(starts, [(0, "外协件"), (2, "外购件"), (4, "外协件")])
            self.assertEqual(starts, category_starts(records))
            with zipfile.ZipFile(path) as archive:
                doc = ET.fromstring(archive.read("word/document.xml"))
            tables = doc.findall("w:body/w:tbl", NS)
            self.assertEqual([sum(category_label(r) is not None for r in t.findall("w:tr", NS))
                              for t in tables], [1, 0, 1, 1])
            for table in tables:
                rows = table.findall("w:tr", NS)
                for i, row in enumerate(rows):
                    if category_label(row):
                        cells = row.findall("w:tc", NS)
                        self.assertEqual([plain_text(c) for c in cells[:3] + cells[4:]], [""] * 8)
                        self.assertIsNotNone(cells[3].find("w:p/w:r/w:rPr/w:b", NS))
                        self.assertEqual(cells[3].find("w:p/w:pPr/w:jc", NS).get(W + "val"), "center")
                        self.assertTrue(all(p.find("w:pPr/w:keepNext", NS) is not None
                                            for p in row.findall("w:tc/w:p", NS)))
                        self.assertTrue(plain_text(rows[i + 1].findall("w:tc", NS)[1]))

    def test_no_pcb_only_has_purchased_heading(self):
        template = Template(RESOURCES / "bom_template.docx")
        record = ["1", "PCB制板配件", "插座", "1", "J1", ""]
        self.assertEqual(material_category(record), "外购件")
        table = template.table([record], True)
        self.assertEqual([category_label(r) for r in table.findall("w:tr", NS) if category_label(r)], ["外购件"])
        continuation = template.table([record], False, previous_category="外购件")
        self.assertFalse(any(category_label(r) for r in continuation.findall("w:tr", NS)))

    def test_measurement_uses_previous_category_and_keeps_category_with_item(self):
        template = Template(RESOURCES / "bom_template.docx")
        record = ["2", "0402", "电阻", "1", "R1", ""]
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "measure.docx"
            template.write(path, [[], [record], [record]], first_page=False, previous_category="外协件")
            with zipfile.ZipFile(path) as archive:
                tables = ET.fromstring(archive.read("word/document.xml")).findall("w:body/w:tbl", NS)
            self.assertEqual([sum(category_label(r) is not None for r in t.findall("w:tr", NS))
                              for t in tables], [0, 1, 0])
            row = next(r for r in tables[1].findall("w:tr", NS) if category_label(r))
            reference = template.data_rows[1].findall("w:tc", NS)
            for cell, data_cell in zip(row.findall("w:tc", NS), reference):
                for tag in ("tcW", "gridSpan"):
                    got = cell.find("w:tcPr/w:" + tag, NS)
                    expected = data_cell.find("w:tcPr/w:" + tag, NS)
                    self.assertEqual(got.attrib if got is not None else None,
                                     expected.attrib if expected is not None else None)

    def test_missing_measured_row_never_silently_drops_material(self):
        template = Template(RESOURCES / "bom_template.docx")
        records = [["1", "0402", "电阻", "1", "R1", ""],
                   ["2", "0402", "电阻", "1", "R2", ""]]
        layout = Mock()
        layout.render.side_effect = [
            {"pages": 1, "tables": [{"start_page": 1, "end_page": 1, "end_y": 500}]},
            {"pages": 2, "tables": [{"start_page": 1, "end_page": 2, "end_y": 500}]},
            {"pages": 1, "tables": [{"start_page": 1, "end_page": 1, "end_y": 100}]},
        ]
        with tempfile.TemporaryDirectory() as folder:
            with self.assertRaisesRegex(StandardError, "表格数量"):
                paginate(template, records, layout, Path(folder))


class ExportTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.input = self.root / "procurement.xlsx"
        self.output = self.root / "standard.docx"
        self.records = [["1", "0402", "电阻", "1", "R1", "10kΩ\n1%"]]
        workbook(self.input, [("BOM", [COLUMNS] + self.records)])

    @staticmethod
    def render(_path, *, pdf=None):
        if pdf is not None:
            pdf.write_bytes(b"%PDF-1.7\nmock Word output")
        return {"pages": 1, "tables": [{"start_page": 1, "end_page": 1, "end_y": 500}]}

    def test_export_verifies_content_and_publishes_all_outputs(self):
        with patch("bom_post_process.standard.check_word_available"), \
             patch("bom_post_process.standard.paginate", return_value=([self.records], [0])), \
             patch("bom_post_process.standard.WordLayout.render", side_effect=self.render):
            result = export_standard_bom(self.input, self.output)
        self.assertEqual((result.items, result.pages, result.sheet), (1, 1, "BOM"))
        self.assertTrue(all(p.exists() for p in (result.output, result.pdf, result.audit)))
        audit = json.loads(result.audit.read_text(encoding="utf-8"))
        self.assertTrue(audit["data_verified"] and audit["layout_verified"] and audit["categories_verified"])
        self.assertEqual(read_exported_content(result.output, Template(RESOURCES / "bom_template.docx"))[0], self.records)

    def test_manual_selection_remarks_are_preserved_without_recalculation(self):
        remarks = ("  人工选型：0603 / 22kΩ / ±0.1% / 0.1W  \n"
                   "替代型号：RC0603FR-0722KL；按最终选型焊接\n\n"
                   "保留字符 & < >，备注数值优先于 Comment\n")
        name = "  人工选择的精密电阻 & < >  "
        records = [["007", "0402-1002F", name, "2", "R1, R2", remarks]]
        workbook(self.input, [("BOM", [COLUMNS] + records)])
        original = self.input.read_bytes()
        with patch("bom_post_process.standard.check_word_available"), \
             patch("bom_post_process.standard.paginate",
                   side_effect=lambda _template, rows, _layout, _work: ([rows], [0])), \
             patch("bom_post_process.standard.WordLayout.render", side_effect=self.render):
            result = export_standard_bom(self.input, self.output)
        exported, _ = read_exported_content(result.output, Template(RESOURCES / "bom_template.docx"))
        self.assertEqual(exported, records)
        self.assertEqual(exported[0][2], name)
        self.assertEqual(exported[0][5], remarks)
        self.assertEqual(self.input.read_bytes(), original)

    def test_procurement_pcb_placeholder_has_both_category_headings(self):
        records = [["1", "PCB制板", "PCB制板", "", "", "PCB制板的占位符"],
                   ["2", "0402K1002", "人工名称", "1", "R1", "人工选型"]]
        # 焊接表可能含未采购器件；Word 应读取 BOM 表中的名称、备注与 PCB 占位。
        welding = [["1", "NOT_BOUGHT", "无需采购", "1", "TP1", "不应导出"]]
        workbook(self.input, [("焊接表", [COLUMNS] + welding), ("BOM", [COLUMNS] + records)])
        with patch("bom_post_process.standard.check_word_available"), \
             patch("bom_post_process.standard.paginate",
                   side_effect=lambda _template, rows, _layout, _work: ([rows], [0])), \
             patch("bom_post_process.standard.WordLayout.render", side_effect=self.render):
            result = export_standard_bom(self.input, self.output)
        exported, categories = read_exported_content(result.output, Template(RESOURCES / "bom_template.docx"))
        self.assertEqual(exported, records)
        self.assertEqual(categories, [(0, "外协件"), (1, "外购件")])
        self.assertEqual((result.items, result.sheet), (2, "BOM"))
        audit = json.loads(result.audit.read_text(encoding="utf-8"))
        self.assertEqual(audit["category_rows"], 2)

    def test_verification_failure_does_not_publish_anything(self):
        with patch("bom_post_process.standard.check_word_available"), \
             patch("bom_post_process.standard.paginate", return_value=([self.records], [0])), \
             patch("bom_post_process.standard.WordLayout.render", side_effect=self.render), \
             patch("bom_post_process.standard.read_exported_content", return_value=([], [])):
            with self.assertRaisesRegex(StandardError, "导出数据与 Excel 不一致"):
                export_standard_bom(self.input, self.output)
        self.assertEqual(list(self.root.iterdir()), [self.input])

    def test_existing_pdf_prevents_export_and_remains_unchanged(self):
        pdf = self.output.with_suffix(".pdf")
        pdf.write_bytes(b"keep existing")
        with patch("bom_post_process.standard.check_word_available") as check:
            with self.assertRaisesRegex(StandardError, "输出文件已存在"):
                export_standard_bom(self.input, self.output)
        check.assert_not_called()
        self.assertEqual(pdf.read_bytes(), b"keep existing")
        self.assertFalse(self.output.exists())

    def test_late_publish_collision_rolls_back_only_new_files(self):
        source = self.root / "candidate"
        source.write_bytes(b"new")
        pdf = self.output.with_suffix(".pdf")
        pdf.write_bytes(b"keep existing")
        with self.assertRaises(FileExistsError):
            _publish([source, source], [self.output, pdf])
        self.assertFalse(self.output.exists())
        self.assertEqual(pdf.read_bytes(), b"keep existing")

    def test_explicit_missing_config_is_reported(self):
        with self.assertRaisesRegex(StandardError, "标准化 BOM 导出失败"):
            export_standard_bom(self.input, self.output, config=self.root / "missing.json")


if __name__ == "__main__":
    unittest.main()
