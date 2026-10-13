"""Export a welding workbook using the supplied, Word-measured BOM template."""
from __future__ import annotations

import copy
from contextlib import ExitStack
from dataclasses import dataclass
import json
import math
import os
from pathlib import Path
import posixpath
import re
import shutil
import subprocess
import tempfile
import xml.etree.ElementTree as ET
import zipfile

from .errors import BomError


COLUMNS = ("序号", "规格型号/代号", "名称", "数量", "项目代号", "备注")
CATEGORIES = ("外协件", "外购件")
NS = {
    "s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}
RESOURCES = Path(__file__).resolve().parent / "resources"
W = "{" + NS["w"] + "}"
ET.register_namespace("w", NS["w"])
ET.register_namespace("r", NS["r"])


class StandardError(BomError):
    """A workbook, template, or Word layout could not be exported safely."""


@dataclass(frozen=True)
class StandardReport:
    output: Path
    pdf: Path
    audit: Path
    pages: int
    items: int
    sheet: str


def element(tag: str, **attrs: str) -> ET.Element:
    return ET.Element(W + tag, {W + k: str(v) for k, v in attrs.items()})


def child(parent: ET.Element, tag: str, **attrs: str) -> ET.Element:
    found = parent.find("w:" + tag, NS)
    if found is None:
        found = element(tag)
        parent.append(found)
    found.attrib.update({W + k: str(v) for k, v in attrs.items()})
    return found


def plain_text(node: ET.Element) -> str:
    return "".join(t.text or "" for t in node.iter(W + "t"))


def material_category(values: list[str]) -> str:
    return "外协件" if any(values[i].strip() == "PCB制板" for i in (1, 2)) else "外购件"


def category_label(row: ET.Element) -> str | None:
    """分类行仅名称列有分类标题；带物料序号等内容的行仍属于物料。"""
    cells = row.findall("w:tc", NS)
    if len(cells) == 9:
        texts = [plain_text(cell).strip() for cell in cells]
        if texts[3] in CATEGORIES and not any(texts[:3] + texts[4:]):
            return texts[3]
    return None


def category_starts(records: list[list[str]]) -> list[tuple[int, str]]:
    starts, previous = [], None
    for index, values in enumerate(records):
        label = material_category(values)
        if label != previous:
            starts.append((index, label))
        previous = label
    return starts


def clean_xml(node: ET.Element) -> ET.Element:
    node = copy.deepcopy(node)
    for item in node.iter():
        for key in list(item.attrib):
            if "rsid" in key or key.startswith("{http://schemas.microsoft.com/") or "Ignorable" in key:
                del item.attrib[key]
        for sub in list(item):
            if sub.tag in {W + "bookmarkStart", W + "bookmarkEnd", W + "proofErr", W + "lastRenderedPageBreak"}:
                item.remove(sub)
    return node


def set_cell_text(cell: ET.Element, text: str) -> None:
    """复用该单元格的段落和字符格式，不继承旧的编号或域。"""
    old_p = cell.find("w:p", NS)
    p_pr = copy.deepcopy(old_p.find("w:pPr", NS)) if old_p is not None else None
    r_pr = old_p.find("w:r/w:rPr", NS) if old_p is not None else None
    if r_pr is None and p_pr is not None:
        r_pr = p_pr.find("w:rPr", NS)
    r_pr = copy.deepcopy(r_pr)
    if p_pr is not None:
        for tag in ("numPr", "keepNext", "pageBreakBefore"):
            for node in p_pr.findall("w:" + tag, NS):
                p_pr.remove(node)
    for node in list(cell):
        if node.tag != W + "tcPr":
            cell.remove(node)
    p = ET.SubElement(cell, W + "p")
    if p_pr is not None:
        p.append(p_pr)
    run = ET.SubElement(p, W + "r")
    if r_pr is not None:
        run.append(r_pr)
    for i, line in enumerate(text.split("\n")):
        if i:
            ET.SubElement(run, W + "br")
        t = ET.SubElement(run, W + "t", {"{http://www.w3.org/XML/1998/namespace}space": "preserve"})
        t.text = line


def small_paragraph(page_break: bool = False) -> ET.Element:
    p = element("p")
    props = child(p, "pPr")
    child(props, "spacing", before="0", after="0", line="20", lineRule="exact")
    child(props, "snapToGrid", val="0")
    child(child(props, "rPr"), "sz", val="2")
    if page_break:
        child(props, "pageBreakBefore")
    return p


class Template:
    def __init__(self, path: Path, replacements: dict[str, str] | None = None):
        with zipfile.ZipFile(path) as archive:
            self.styles = clean_xml(ET.fromstring(archive.read("word/styles.xml")))
            doc = ET.fromstring(archive.read("word/document.xml"))
        body = doc.find("w:body", NS)
        if body is None or body.find("w:sectPr", NS) is None:
            raise StandardError("Word 模板缺少正文或页面设置。")
        tables = body.findall("w:tbl", NS)
        if len(tables) < 2:
            raise StandardError("模板至少应包含首页表格和续页表格。")
        self.section = clean_xml(body.find("w:sectPr", NS))
        for tag, attrs in (("pgSz", ("h",)), ("pgMar", ("bottom",))):
            setting = self.section.find("w:" + tag, NS)
            if setting is None or any(setting.get(W + attr) is None for attr in attrs):
                raise StandardError("Word 模板缺少有效的纸张尺寸或页边距。")
        self.tables = [clean_xml(t) for t in tables[:2]]
        self.heads, self.data_rows, self.tails = [], [], []
        self.category_rows = {}
        for table in self.tables:
            rows = table.findall("w:tr", NS)
            header_index = next((i for i, row in enumerate(rows) if "规格型号" in plain_text(row)), None)
            tail_index = next((i for i, row in enumerate(rows) if "PCB订单号" in plain_text(row)), None)
            if header_index is None or tail_index is None or tail_index <= header_index + 1:
                raise StandardError("模板未找到表头、样例数据行或含 PCB订单号 的表尾。")
            body_rows = rows[header_index + 1:tail_index]
            for row in body_rows:
                label = category_label(row)
                if label:
                    self.category_rows.setdefault(label, row)
            # 分类行不能成为物料样式；标准模板中的空白数据行仍可作为样例。
            samples = [r for r in body_rows if len(r.findall("w:tc", NS)) == 9 and not category_label(r)]
            if not samples:
                raise StandardError("模板未找到物料样例行，分类标题不能代替数据行。")
            sample = next((r for r in samples
                           if plain_text(r.findall("w:tc", NS)[2]).strip()
                           and plain_text(r.findall("w:tc", NS)[5]).strip()), samples[0])
            self.heads.append(rows[:header_index + 1])
            self.data_rows.append(sample)
            self.tails.append(rows[tail_index:])
            for row in rows:
                # 侧边留白无需跨整页纵向合并，拆开后视觉一致，避免旧 DOC 的分页死循环。
                cells = row.findall("w:tc", NS)
                if not cells:
                    raise StandardError("Word 模板包含没有单元格的空行。")
                for cell in (cells[0], cells[-1]):
                    props = child(cell, "tcPr")
                    merge = props.find("w:vMerge", NS)
                    if merge is not None:
                        props.remove(merge)
                child(child(row, "trPr"), "cantSplit")
            for row in self.tails[-1]:
                for cell in row.findall("w:tc", NS):
                    label = plain_text(cell).strip()
                    if replacements and label in replacements:
                        set_cell_text(cell, replacements[label])
        if any(label not in self.category_rows for label in CATEGORIES):
            raise StandardError("模板缺少外协件或外购件分类行，请使用随项目提供的 bom_template.docx。")

    def category_row(self, label: str, kind: int) -> ET.Element:
        row = copy.deepcopy(self.category_rows[label])
        # 分类行沿用示例的样式，网格跨度和列宽适配首页或续页。
        for cell, reference in zip(row.findall("w:tc", NS), self.data_rows[kind].findall("w:tc", NS)):
            props = child(cell, "tcPr")
            for tag in ("tcW", "gridSpan", "vMerge"):
                for node in props.findall("w:" + tag, NS):
                    props.remove(node)
            geometry = child(reference, "tcPr")
            for index, tag in enumerate(("tcW", "gridSpan")):
                value = geometry.find("w:" + tag, NS)
                if value is not None:
                    props.insert(index, copy.deepcopy(value))
            for paragraph in cell.findall("w:p", NS):
                child(child(paragraph, "pPr"), "keepNext")
        return row

    def table(self, records: list[list[str]], first: bool, page: int = 1,
              total: int = 1, filler: int = 0, *, previous_category: str | None = None) -> ET.Element:
        kind = 0 if first else 1
        table = copy.deepcopy(self.tables[kind])
        for row in table.findall("w:tr", NS):
            table.remove(row)
        for row in self.heads[kind]:
            table.append(copy.deepcopy(row))
        for values in records:
            label = material_category(values)
            if label != previous_category:
                table.append(self.category_row(label, kind))
            previous_category = label
            row = copy.deepcopy(self.data_rows[kind])
            for cell, value in zip(row.findall("w:tc", NS)[1:7], values):
                set_cell_text(cell, value)
            set_cell_text(row.findall("w:tc", NS)[7], "")
            # 序号列原为 Word 自动编号；改为 Excel 序号后，保留列宽并容纳三位数。
            serial = row.findall("w:tc", NS)[1]
            margins = child(child(serial, "tcPr"), "tcMar")
            child(margins, "left", w="20", type="dxa")
            child(margins, "right", w="20", type="dxa")
            child(child(serial.find("w:p", NS), "pPr"), "jc", val="center")
            table.append(row)
        if filler >= 40:
            row = copy.deepcopy(self.data_rows[kind])
            for cell in row.findall("w:tc", NS):
                set_cell_text(cell, "")
                for p in cell.findall("w:p", NS):
                    props = child(p, "pPr")
                    child(props, "spacing", before="0", after="0", line="20", lineRule="exact")
                    child(props, "snapToGrid", val="0")
            child(child(row, "trPr"), "trHeight", val=str(filler), hRule="exact")
            table.append(row)
        for original in self.tails[kind]:
            row = copy.deepcopy(original)
            for cell in row.findall("w:tc", NS):
                text = plain_text(cell)
                if "共" in text and "张" in text and "第" in text:
                    set_cell_text(cell, f"共 {total} 张    第 {page} 张")
            table.append(row)
        return table

    def write(self, path: Path, groups: list[list[list[str]]], *,
              first_page: bool = True, fillers: list[int] | None = None,
              previous_category: str | None = None) -> None:
        doc = element("document")
        body = child(doc, "body")
        for i, group in enumerate(groups):
            body.append(small_paragraph(page_break=i > 0))
            body.append(self.table(group, first_page and i == 0, i + 1, len(groups),
                                   fillers[i] if fillers else 0, previous_category=previous_category))
            if group:
                previous_category = material_category(group[-1])
        body.append(small_paragraph())
        body.append(copy.deepcopy(self.section))
        settings = element("settings")
        child(settings, "view", val="print")
        compat = child(settings, "compat")
        child(compat, "compatSetting", name="compatibilityMode", uri="http://schemas.microsoft.com/office/word", val="15")
        # 保留模板的中英文换行、字符压缩和表内网格选项；不沿用旧 DOC 布局模式。
        child(compat, "doNotSnapToGridInCell")
        child(settings, "characterSpacingControl", val="compressPunctuation")
        parts = {"word/document.xml": doc, "word/styles.xml": self.styles, "word/settings.xml": settings}
        types_ns = "http://schemas.openxmlformats.org/package/2006/content-types"
        rel_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
        types = ET.Element("Types", xmlns=types_ns)
        ET.SubElement(types, "Default", Extension="rels", ContentType="application/vnd.openxmlformats-package.relationships+xml")
        for name, suffix in (("document", "document.main"), ("styles", "styles"), ("settings", "settings")):
            ET.SubElement(types, "Override", PartName=f"/word/{name}.xml",
                          ContentType=f"application/vnd.openxmlformats-officedocument.wordprocessingml.{suffix}+xml")
        relations = ET.Element("Relationships", xmlns=rel_ns)
        ET.SubElement(relations, "Relationship", Id="rId1", Type=NS["r"] + "/officeDocument", Target="word/document.xml")
        doc_relations = ET.Element("Relationships", xmlns=rel_ns)
        for i, name in enumerate(("styles", "settings"), 1):
            ET.SubElement(doc_relations, "Relationship", Id=f"rId{i}", Type=NS["r"] + "/" + name, Target=name + ".xml")
        parts.update({"[Content_Types].xml": types, "_rels/.rels": relations, "word/_rels/document.xml.rels": doc_relations})
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
            for name, tree in parts.items():
                archive.writestr(name, ET.tostring(tree, encoding="utf-8", xml_declaration=True))


def read_bom(path: Path, sheet_name: str | None = None) -> tuple[str, list[list[str]]]:
    """Read six named columns; prefer 焊接表, otherwise require one matching sheet."""
    with zipfile.ZipFile(path) as archive:
        workbook = ET.fromstring(archive.read("xl/workbook.xml"))
        sheets = workbook.findall("s:sheets/s:sheet", NS)
        names = [s.get("name") for s in sheets]
        if sheet_name and sheet_name not in names:
            raise StandardError(f"找不到工作表 {sheet_name!r}；可用：{names}")
        selected = sheet_name or ("焊接表" if "焊接表" in names else None)
        relations = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        targets = {r.get("Id"): r.get("Target") for r in relations}
        strings = []
        if "xl/sharedStrings.xml" in archive.namelist():
            strings = ["".join(t.text or "" for t in item.findall(".//s:t", NS))
                       for item in ET.fromstring(archive.read("xl/sharedStrings.xml"))]

        def cell_text(cell: ET.Element, *, strict: bool = True) -> str:
            value = cell.find("s:v", NS)
            raw = value.text or "" if value is not None else ""
            kind = cell.get("t")
            if strict and kind == "e":
                raise StandardError(f"Excel 单元格 {cell.get('r')} 含错误：{raw}")
            if strict and cell.find("s:f", NS) is not None and not raw:
                raise StandardError(f"公式 {cell.get('r')} 没有计算结果，请在 Excel 中重算并保存。")
            if kind == "s":
                return strings[int(raw)]
            if kind == "inlineStr":
                return "".join(t.text or "" for t in cell.findall(".//s:t", NS))
            if kind == "b":
                return "TRUE" if raw == "1" else "FALSE"
            return raw

        candidates = []
        for sheet in sheets:
            if selected and sheet.get("name") != selected:
                continue
            target = targets[sheet.get(f"{{{NS['r']}}}id")]
            sheet_path = target.lstrip("/") if target.startswith("/") else posixpath.normpath("xl/" + target)
            rows = ET.fromstring(archive.read(sheet_path)).findall("s:sheetData/s:row", NS)
            for index, row in enumerate(rows):
                cells = {re.sub(r"\d", "", c.get("r", "")): c for c in row.findall("s:c", NS)}
                values = {col: cell_text(c, strict=False).strip() for col, c in cells.items()}
                if all(label in values.values() for label in COLUMNS):
                    if any(list(values.values()).count(label) != 1 for label in COLUMNS):
                        raise StandardError(f"工作表 {sheet.get('name')} 必需列名重复，请每个字段只保留一列。")
                    headers = [next(col for col, val in values.items() if val == label) for label in COLUMNS]
                    candidates.append((sheet.get("name"), headers, rows[index + 1:]))
                    break
        if not candidates:
            raise StandardError("Excel 缺少表头（焊接表），必需列：" + "、".join(COLUMNS)
                                + "；请先运行 gen_value，完成人工选型后再运行 gen_soldering。")
        if len(candidates) > 1:
            raise StandardError("多个工作表符合 BOM 表头，请使用 --sheet 指定："
                                + "、".join(candidate[0] for candidate in candidates))
        name, headers, rows = candidates[0]
        result = []
        for row in rows:
            cells = {re.sub(r"\d", "", c.get("r", "")): c for c in row.findall("s:c", NS)}
            values = [cell_text(cells[col]) if col in cells else "" for col in headers]
            if any(v.strip() for v in values):
                result.append([v.replace("\r\n", "\n").replace("\r", "\n") for v in values])
        if not result:
            raise StandardError("Excel 中没有 BOM 数据。")
        return name, result


def windows_path(path: Path) -> str:
    return str(path.resolve())


def check_word_available() -> None:
    """Check the required Windows/Word installation before export."""
    if os.name != "nt":
        raise StandardError("Word 标准化需要 Windows 和已安装的桌面版 Microsoft Word；validate、gen_value 和 gen_soldering 可独立使用。")
    if not shutil.which("powershell.exe"):
        raise StandardError("未找到 powershell.exe；Word 标准化需要 Windows PowerShell。")
    import winreg
    try:
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"Word.Application\CLSID"):
            pass
    except OSError as exc:
        raise StandardError("未检测到桌面版 Microsoft Word；请安装并激活 Word 后运行 gen_bom_standard。") from exc


def read_exported_content(path: Path, template: Template) -> tuple[list[list[str]], list[tuple[int, str]]]:
    """回读实际正文，分别核验物料数据和分类位置，忽略表尾及补高空行。"""
    with zipfile.ZipFile(path) as archive:
        doc = ET.fromstring(archive.read("word/document.xml"))
    records, categories = [], []
    for i, table in enumerate(doc.findall("w:body/w:tbl", NS)):
        kind = 0 if i == 0 else 1
        rows = table.findall("w:tr", NS)[len(template.heads[kind]):-len(template.tails[kind])]
        for index, row in enumerate(rows):
            label = category_label(row)
            if label:
                following = rows[index + 1].findall("w:tc", NS) if index + 1 < len(rows) else []
                if (len(following) != 9 or category_label(rows[index + 1])
                        or not any(plain_text(c).strip() for c in following[1:7])):
                    raise StandardError("分类标题未与首条物料放在同一页，未发布文件。")
                categories.append((len(records), label))
                continue
            cells = row.findall("w:tc", NS)
            if len(cells) != 9:
                raise StandardError("导出正文列数异常，未发布文件。")
            values = ["".join("\n" if n.tag == W + "br" else (n.text or "")
                              for n in cell.iter() if n.tag in (W + "t", W + "br")) for cell in cells[1:7]]
            if any(v.strip() for v in values):
                records.append(values)
    return records, categories


class WordLayout:
    def __init__(self, directory: Path):
        self.directory = directory
        self.counter = 0

    def render(self, path: Path, *, pdf: Path | None = None) -> dict:
        self.counter += 1
        job_path = self.directory / f"job_{self.counter}.json"
        result_path = self.directory / f"layout_{self.counter}.json"
        job = {"input": windows_path(path), "result": windows_path(result_path)}
        if pdf:
            job["pdf"] = windows_path(pdf)
        job_path.write_text(json.dumps(job, ensure_ascii=True), encoding="utf-8")
        cmd = ["powershell.exe", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass",
               "-File", windows_path(RESOURCES / "word_layout.ps1"), "-JobFile", windows_path(job_path)]
        try:
            result = subprocess.run(cmd, capture_output=True, encoding="utf-8", errors="replace", timeout=180,
                                    creationflags=subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0)
        except subprocess.TimeoutExpired as exc:
            raise StandardError("Word 排版超时（180 秒）；请检查 Word 是否有等待处理的对话框。") from exc
        if result.returncode != 0:
            raise StandardError("Word 排版失败；请确认桌面版 Microsoft Word 可正常启动且已激活："
                                + (result.stderr or result.stdout).strip())
        if not result_path.exists():
            raise StandardError("Word 未返回排版结果。请确认已安装并激活桌面版 Microsoft Word。")
        result = json.loads(result_path.read_text(encoding="utf-8-sig"))
        _check_layout(result)
        return result


def _check_layout(result: dict, expected_tables: int | None = None) -> None:
    if not isinstance(result, dict) or not isinstance(result.get("pages"), int) or result["pages"] < 1:
        raise StandardError("Word 未返回有效的页数。")
    tables = result.get("tables")
    if not isinstance(tables, list) or not tables or (expected_tables is not None and len(tables) != expected_tables):
        raise StandardError("Word 返回的表格数量与试排文档不一致，未发布文件。")
    for table in tables:
        if (not isinstance(table, dict)
                or any(not isinstance(table.get(key), (int, float)) or not math.isfinite(table[key])
                       for key in ("start_page", "end_page", "end_y"))
                or not 1 <= table["start_page"] <= table["end_page"] <= result["pages"]
                or table["end_y"] < 0):
            raise StandardError("Word 返回了无效的表格页码或行高，未发布文件。")


def paginate(template: Template, records: list[list[str]], layout: WordLayout, work: Path) -> tuple[list, list]:
    if not records:
        raise StandardError("Excel 中没有 BOM 数据。")
    size = template.section.find("w:pgSz", NS)
    margins = template.section.find("w:pgMar", NS)
    # 末尾留白行和文末段落也占高度；额外保留 20 pt，避免打印布局中的边界舍入。
    target_y = (int(size.get(W + "h")) - int(margins.get(W + "bottom"))) / 20 - 20
    probe = work / "probe.docx"
    # 首页列宽和表尾均不同，单独用 Word 二分试排。
    minimum = template.data_rows[0].find("w:trPr/w:trHeight", NS)
    minimum_height = int(minimum.get(W + "val", "20")) / 20 if minimum is not None else 1
    low, high = 0, min(len(records), int(target_y / max(minimum_height, 1)))
    while low < high:
        count = (low + high + 1) // 2
        template.write(probe, [records[:count]])
        result = layout.render(probe)
        _check_layout(result, 1)
        fits = result["pages"] == 1 and result["tables"][0]["end_y"] <= target_y
        if fits:
            low = count
        else:
            high = count - 1
    if not low:
        raise StandardError(f"序号 {records[0][0]} 的内容无法放入首页；请缩短该行内容或调整模板列宽。")
    groups = [records[:low]]
    print(f"首页放入 {low} 条物料；正在测量续页行高……", flush=True)
    remaining = records[low:]
    if remaining:
        # 分类标题与其第一条物料一起测量；承接首页分类，续页不重复同类标题。
        measure = work / "measure.docx"
        template.write(measure, [[]] + [[row] for row in remaining], first_page=False,
                       previous_category=material_category(records[low - 1]))
        measured = layout.render(measure)
        _check_layout(measured, len(remaining) + 1)
        base = measured["tables"][0]["end_y"]
        heights = []
        for row, table in zip(remaining, measured["tables"][1:]):
            if table["start_page"] != table["end_page"] or table["end_y"] > target_y:
                raise StandardError(f"序号 {row[0]} 的一整行超出单页可用高度，无法在保留字号且不拆行的条件下导出。")
            if table["end_y"] < base:
                raise StandardError("Word 测得的物料行高为负，未发布文件。")
            heights.append(table["end_y"] - base)
        current, used = [], base
        for row, height in zip(remaining, heights):
            if current and used + height > target_y:
                groups.append(current)
                current, used = [], base
            current.append(row)
            used += height
        if current:
            groups.append(current)
    # 在完整文档中重新实测。表格相邻边框的合并可能造成小量差异，溢出则把末行顺延。
    for attempt in range(len(records) + 1):
        template.write(probe, groups)
        measured = layout.render(probe)
        _check_layout(measured, len(groups))
        changed = False
        for i, table in enumerate(measured["tables"]):
            if table["start_page"] != table["end_page"] or table["end_y"] > target_y:
                if len(groups[i]) < 2:
                    raise StandardError(f"序号 {groups[i][0][0]} 无法完整排入一页。")
                last = groups[i].pop()
                if i + 1 == len(groups):
                    groups.append([])
                groups[i + 1].insert(0, last)
                changed = True
        if not changed:
            if measured["pages"] != len(groups):
                raise StandardError("Word 出现额外空白页，未发布导出文件。")
            # 把表尾对齐到固定底部位置。预留 2 pt 给新插入行的边框。
            fillers = [max(0, int((target_y - t["end_y"] - 2) * 20)) for t in measured["tables"]]
            return groups, fillers
    raise StandardError("分页未收敛，未发布导出文件。")


def _publish(sources: list[Path], destinations: list[Path]) -> None:
    """Reserve every final file exclusively, and roll back this operation on failure."""
    created = []
    try:
        with ExitStack() as stack:
            handles = []
            for destination in destinations:
                handle = stack.enter_context(destination.open("xb"))
                created.append(destination)
                handles.append(handle)
            for source, handle in zip(sources, handles):
                with source.open("rb") as stream:
                    shutil.copyfileobj(stream, handle)
    except BaseException:
        for path in created:
            path.unlink(missing_ok=True)
        raise


def export_standard_bom(
    input_path: str | Path,
    output_path: str | Path,
    *,
    sheet: str | None = None,
    template: str | Path | None = None,
    config: str | Path | None = None,
) -> StandardReport:
    """Generate DOCX, PDF and an audit JSON after Word layout and content checks.

    Existing files are never overwritten.  All final files are published only after
    successful verification; a failed publish removes the files created by this call.
    """
    try:
        return _export_standard_bom(Path(input_path), Path(output_path), sheet=sheet,
                                    template=Path(template) if template is not None else RESOURCES / "bom_template.docx",
                                    config=Path(config) if config is not None else RESOURCES / "bom_config.json")
    except StandardError:
        raise
    except (OSError, ValueError, KeyError, IndexError, ET.ParseError, zipfile.BadZipFile) as exc:
        raise StandardError(f"标准化 BOM 导出失败：{exc}") from exc


def _export_standard_bom(input_path: Path, output: Path, *, sheet: str | None,
                         template: Path, config: Path) -> StandardReport:
    if input_path.suffix.lower() != ".xlsx":
        raise StandardError("仅支持 .xlsx，请先在 Excel 中另存为此格式。")
    output = output.resolve()
    if output.suffix.lower() != ".docx":
        raise StandardError("Word 输出路径必须以 .docx 结尾。")
    outputs = [output, output.with_suffix(".pdf"), output.with_suffix(".json")]
    for path in outputs:
        if path.exists():
            raise StandardError(f"输出文件已存在：{path}；请指定新名称，原文件不会被覆盖。")
    sheet_name, records = read_bom(input_path, sheet)
    replacements = json.loads(config.read_text(encoding="utf-8-sig"))
    if not isinstance(replacements, dict):
        raise StandardError("表尾配置必须是 JSON 对象。")
    replacements = {key: str(value) for key, value in replacements.items() if value not in (None, "")}
    template = Template(template, replacements)
    check_word_available()
    output.parent.mkdir(parents=True, exist_ok=True)
    print(f"读取 {sheet_name}：{len(records)} 条物料。", flush=True)
    # Keep all probes in a temporary subdirectory so failed runs leave no final files.
    with tempfile.TemporaryDirectory(prefix=".bom_", dir=output.parent) as directory:
        work = Path(directory)
        layout = WordLayout(work)
        groups, fillers = paginate(template, records, layout, work)
        candidate, pdf = work / "final.docx", work / "final.pdf"
        template.write(candidate, groups, fillers=fillers)
        print(f"共 {len(groups)} 页，正在逐页核验并生成 PDF……", flush=True)
        result = layout.render(candidate, pdf=pdf)
        _check_layout(result, len(groups))
        if result["pages"] != len(groups) or len(result["tables"]) != len(groups) or any(
            t["start_page"] != i or t["end_page"] != i
            for i, t in enumerate(result["tables"], 1)
        ):
            raise StandardError("最终分页核验失败，未发布文件；表头、正文和表尾必须全部位于同一页。")
        exported, categories = read_exported_content(candidate, template)
        if exported != records:
            raise StandardError("导出数据与 Excel 不一致，未发布文件。")
        if categories != category_starts(records):
            raise StandardError("导出分类行的位置或内容不正确，未发布文件。")
        if not pdf.is_file() or not pdf.stat().st_size:
            raise StandardError("Word 未生成有效的 PDF，未发布文件。")
        report = {"excel": str(input_path.resolve()), "sheet": sheet_name, "records": len(records),
                  "pages": len(groups), "data_verified": True, "layout_verified": True,
                  "categories_verified": True, "category_rows": len(categories),
                  "page_records": [{"page": i, "count": len(group), "first": group[0][0], "last": group[-1][0]}
                                   for i, group in enumerate(groups, 1)], "word_layout": result}
        audit = work / "final.json"
        audit.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
        _publish([candidate, pdf, audit], outputs)
    return StandardReport(*outputs, pages=len(groups), items=len(records), sheet=sheet_name)
